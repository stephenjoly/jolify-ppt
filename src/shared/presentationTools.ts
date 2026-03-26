import JSZip from "jszip";

import type { ActionResult } from "./shapeTools";

const REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const LOCAL_BRIDGE_ORIGIN = "https://127.0.0.1:38443";
const PptRelationship = {
  comments: "/relationships/comments",
  notesSlide: "/relationships/notesSlide",
  slide: "/relationships/slide",
  commentAuthors: "/relationships/commentAuthors",
  slideMaster: "/relationships/slideMaster",
  theme: "/relationships/theme",
} as const;

type SlideExportMode = "download" | "local-save";

type SlidePartInfo = {
  id: string;
  index: number;
  relId: string;
  target: string;
  title: string;
  commentsTarget?: string;
  notesTarget?: string;
};

type PresentationStructure = {
  slides: SlidePartInfo[];
  presentationDoc: XMLDocument;
  presentationRelsDoc: XMLDocument;
};

export type ThemePalette = {
  source: "deck" | "fallback";
  colors: string[];
  rowSize: number;
};

function parseXml(xml: string): XMLDocument {
  return new DOMParser().parseFromString(xml, "application/xml");
}

function serializeXml(doc: XMLDocument): string {
  return new XMLSerializer().serializeToString(doc);
}

function localNameOf(node: Node | null): string {
  return node?.localName ?? "";
}

function childElements(element: Element, tagName?: string): Element[] {
  return Array.from(element.childNodes).filter((node): node is Element => {
    return node.nodeType === Node.ELEMENT_NODE && (!tagName || localNameOf(node) === tagName);
  });
}

function toPosixPath(path: string): string {
  const parts = path.split("/").filter((part) => part && part !== ".");
  const resolved: string[] = [];

  parts.forEach((part) => {
    if (part === "..") {
      resolved.pop();
      return;
    }
    resolved.push(part);
  });

  return resolved.join("/");
}

function dirname(path: string): string {
  const idx = path.lastIndexOf("/");
  return idx === -1 ? "" : path.slice(0, idx);
}

function resolvePartPath(basePath: string, target: string): string {
  const baseDir = dirname(basePath);
  return toPosixPath(`${baseDir}/${target}`);
}

function fileNameWithoutExtension(path: string): string {
  const name = path.split("/").pop() ?? path;
  return name.replace(/\.[^.]+$/, "");
}

function getRelationshipId(element: Element): string {
  return element.getAttributeNS(REL_NS, "id") ?? element.getAttribute("r:id") ?? "";
}

function collectText(element: Element): string[] {
  const texts: string[] = [];
  const walker = element.ownerDocument.createTreeWalker(element, NodeFilter.SHOW_ELEMENT | NodeFilter.SHOW_TEXT);

  let node: Node | null = walker.currentNode;
  while (node) {
    if (node.nodeType === Node.TEXT_NODE) {
      const text = node.textContent?.trim();
      if (text) {
        texts.push(text);
      }
    }
    node = walker.nextNode();
  }

  return texts;
}

function getFirstMeaningfulText(element: Element): string {
  const texts = collectText(element);
  return texts.find((text) => text.trim() !== "") ?? "";
}

async function readZipText(zip: JSZip, path: string): Promise<string | null> {
  const file = zip.file(path);
  return file ? file.async("string") : null;
}

function toBase64(bytes: Uint8Array): string {
  let binary = "";
  const chunk = 0x8000;
  for (let i = 0; i < bytes.length; i += chunk) {
    binary += String.fromCharCode(...bytes.subarray(i, i + chunk));
  }
  return btoa(binary);
}

function getBaseNameFromUrl(url: string | null | undefined): string {
  if (!url) {
    return "Presentation";
  }

  try {
    const pathname = new URL(url).pathname;
    const name = pathname.split("/").pop() ?? "Presentation";
    return decodeURIComponent(name).replace(/\.pptx$/i, "") || "Presentation";
  } catch {
    return url.replace(/^.*[\\/]/, "").replace(/\.pptx$/i, "") || "Presentation";
  }
}

function downloadFile(filename: string, mimeType: string, content: BlobPart) {
  const blob = new Blob([content], { type: mimeType });
  const url = URL.createObjectURL(blob);
  const anchor = document.createElement("a");
  anchor.href = url;
  anchor.download = filename;
  anchor.click();
  setTimeout(() => URL.revokeObjectURL(url), 1000);
}

function sanitizeFilename(filename: string): string {
  return filename.replace(/[/:*?"<>|]/g, "-").trim() || "Jolify Export";
}

function fileToBase64(file: File): Promise<string> {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onerror = () => reject(new Error(`Could not read ${file.name}.`));
    reader.onload = () => {
      const result = reader.result;
      if (typeof result !== "string") {
        reject(new Error(`Could not read ${file.name}.`));
        return;
      }

      const commaIndex = result.indexOf(",");
      resolve(commaIndex === -1 ? result : result.slice(commaIndex + 1));
    };
    reader.readAsDataURL(file);
  });
}

function buildMarkdownSection(title: string, body: string): string {
  return `## ${title}\n\n${body.trim()}\n`;
}

function hexToRgb(hex: string): { r: number; g: number; b: number } | null {
  const normalized = hex.replace("#", "").trim();
  if (!/^[0-9a-fA-F]{6}$/.test(normalized)) {
    return null;
  }

  return {
    r: parseInt(normalized.slice(0, 2), 16),
    g: parseInt(normalized.slice(2, 4), 16),
    b: parseInt(normalized.slice(4, 6), 16),
  };
}

function rgbToHex(r: number, g: number, b: number): string {
  const clamp = (value: number) => Math.max(0, Math.min(255, Math.round(value)));
  return `#${[clamp(r), clamp(g), clamp(b)].map((value) => value.toString(16).padStart(2, "0")).join("").toUpperCase()}`;
}

function mixHex(baseHex: string, targetHex: string, amount: number): string | null {
  const base = hexToRgb(baseHex);
  const target = hexToRgb(targetHex);
  if (!base || !target) {
    return null;
  }

  return rgbToHex(
    base.r + (target.r - base.r) * amount,
    base.g + (target.g - base.g) * amount,
    base.b + (target.b - base.b) * amount,
  );
}

function deriveThemeScale(baseHex: string): string[] {
  const top = baseHex.toUpperCase();
  const row2 = mixHex(baseHex, "#FFFFFF", 0.68);
  const row3 = mixHex(baseHex, "#FFFFFF", 0.38);
  const row4 = mixHex(baseHex, "#000000", 0.22);
  const row5 = mixHex(baseHex, "#000000", 0.46);

  return [top, row2, row3, row4, row5].filter((color): color is string => !!color);
}

function flattenThemeRows(scales: string[][]): string[] {
  const rowCount = Math.max(...scales.map((scale) => scale.length), 0);
  const ordered: string[] = [];

  for (let rowIndex = 0; rowIndex < rowCount; rowIndex += 1) {
    scales.forEach((scale) => {
      const color = scale[rowIndex];
      if (color) {
        ordered.push(color);
      }
    });
  }

  return ordered;
}

function readFileSlice(file: Office.File, index: number): Promise<Uint8Array> {
  return new Promise((resolve, reject) => {
    file.getSliceAsync(index, (result) => {
      if (result.status === Office.AsyncResultStatus.Failed) {
        reject(result.error);
        return;
      }

      const data = result.value.data as number[] | ArrayBuffer;
      if (data instanceof ArrayBuffer) {
        resolve(new Uint8Array(data));
        return;
      }

      resolve(new Uint8Array(data));
    });
  });
}

async function getDocumentBytes(fileType: Office.FileType = Office.FileType.Compressed): Promise<Uint8Array> {
  const file = await new Promise<Office.File>((resolve, reject) => {
    Office.context.document.getFileAsync(fileType, { sliceSize: 65536 }, (result) => {
      if (result.status === Office.AsyncResultStatus.Failed) {
        reject(result.error);
        return;
      }
      resolve(result.value);
    });
  });

  try {
    const slices = await Promise.all(
      Array.from({ length: file.sliceCount }, (_, index) => readFileSlice(file, index)),
    );

    const totalSize = slices.reduce((sum, slice) => sum + slice.length, 0);
    const bytes = new Uint8Array(totalSize);
    let offset = 0;
    slices.forEach((slice) => {
      bytes.set(slice, offset);
      offset += slice.length;
    });
    return bytes;
  } finally {
    file.closeAsync();
  }
}

async function getFileProperties(): Promise<Office.FileProperties> {
  return new Promise((resolve, reject) => {
    Office.context.document.getFilePropertiesAsync((result) => {
      if (result.status === Office.AsyncResultStatus.Failed) {
        reject(result.error);
        return;
      }
      resolve(result.value);
    });
  });
}

async function getPresentationStructure(zip: JSZip): Promise<PresentationStructure> {
  const presentationXml = await readZipText(zip, "ppt/presentation.xml");
  const presentationRelsXml = await readZipText(zip, "ppt/_rels/presentation.xml.rels");

  if (!presentationXml || !presentationRelsXml) {
    throw new Error("The presentation package is missing its core slide index.");
  }

  const presentationDoc = parseXml(presentationXml);
  const presentationRelsDoc = parseXml(presentationRelsXml);
  const relMap = new Map<string, { target: string; type: string }>();

  Array.from(presentationRelsDoc.getElementsByTagName("*")).forEach((node) => {
    if (localNameOf(node) !== "Relationship") {
      return;
    }

    const id = node.getAttribute("Id");
    const target = node.getAttribute("Target");
    const type = node.getAttribute("Type");
    if (id && target && type) {
      relMap.set(id, { target: resolvePartPath("ppt/presentation.xml", target), type });
    }
  });

  const slideNodes = Array.from(presentationDoc.getElementsByTagName("*")).filter(
    (node): node is Element => localNameOf(node) === "sldId",
  );

  const slides = await Promise.all(
    slideNodes.map(async (node, index) => {
      const relId = getRelationshipId(node);
      const relInfo = relMap.get(relId);
      if (!relInfo) {
        throw new Error(`Missing slide relationship ${relId}.`);
      }

      const target = relInfo.target;
      const slideXml = await readZipText(zip, target);
      if (!slideXml) {
        throw new Error(`Missing slide part ${target}.`);
      }

      const slideDoc = parseXml(slideXml);
      const title = getFirstMeaningfulText(slideDoc.documentElement) || `Slide ${index + 1}`;
      const slideRelsPath = `${dirname(target)}/_rels/${fileNameWithoutExtension(target)}.xml.rels`;
      const slideRelsXml = await readZipText(zip, slideRelsPath);

      let commentsTarget: string | undefined;
      let notesTarget: string | undefined;

      if (slideRelsXml) {
        const slideRelsDoc = parseXml(slideRelsXml);
        Array.from(slideRelsDoc.getElementsByTagName("*")).forEach((relNode) => {
          if (localNameOf(relNode) !== "Relationship") {
            return;
          }

          const relType = relNode.getAttribute("Type") ?? "";
          const relTarget = relNode.getAttribute("Target");
          if (!relTarget) {
            return;
          }

          const resolvedTarget = resolvePartPath(target, relTarget);
          if (relType.endsWith(PptRelationship.comments)) {
            commentsTarget = resolvedTarget;
          }
          if (relType.endsWith(PptRelationship.notesSlide)) {
            notesTarget = resolvedTarget;
          }
        });
      }

      return {
        id: node.getAttribute("id") ?? String(index + 1),
        index: index + 1,
        relId,
        target,
        title,
        commentsTarget,
        notesTarget,
      };
    }),
  );

  return { slides, presentationDoc, presentationRelsDoc };
}

async function getPresentationDocuments(zip: JSZip): Promise<{ presentationDoc: XMLDocument; presentationRelsDoc: XMLDocument }> {
  const presentationXml = await readZipText(zip, "ppt/presentation.xml");
  const presentationRelsXml = await readZipText(zip, "ppt/_rels/presentation.xml.rels");

  if (!presentationXml || !presentationRelsXml) {
    throw new Error("The presentation package is missing its core slide index.");
  }

  return {
    presentationDoc: parseXml(presentationXml),
    presentationRelsDoc: parseXml(presentationRelsXml),
  };
}

function findRelationshipTargetByType(
  relsDoc: XMLDocument,
  basePath: string,
  relationshipSuffix: string,
): string | null {
  for (const node of Array.from(relsDoc.getElementsByTagName("*"))) {
    if (localNameOf(node) !== "Relationship") {
      continue;
    }

    const type = node.getAttribute("Type") ?? "";
    const target = node.getAttribute("Target");
    if (type.endsWith(relationshipSuffix) && target) {
      return resolvePartPath(basePath, target);
    }
  }

  return null;
}

function extractThemeSchemeColors(themeDoc: XMLDocument): string[] {
  const colorOrder = ["dk1", "lt1", "dk2", "lt2", "accent1", "accent2", "accent3", "accent4", "accent5", "accent6"];
  const schemeNode = Array.from(themeDoc.getElementsByTagName("*")).find((node): node is Element => localNameOf(node) === "clrScheme");
  if (!schemeNode) {
    return [];
  }

  const colors: string[] = [];

  colorOrder.forEach((name) => {
    const colorNode = childElements(schemeNode).find((child) => localNameOf(child) === name);
    if (!colorNode) {
      return;
    }

    const valueNode = childElements(colorNode).find((child) => {
      const local = localNameOf(child);
      return local === "srgbClr" || local === "sysClr";
    });

    if (!valueNode) {
      return;
    }

    const value =
      valueNode.getAttribute("val") ??
      valueNode.getAttribute("lastClr") ??
      valueNode.getAttribute("lastColor");

    if (value && /^[0-9a-fA-F]{6}$/.test(value)) {
      colors.push(`#${value.toUpperCase()}`);
    }
  });

  return colors;
}

export async function getCurrentPresentationThemePalette(): Promise<ThemePalette> {
  const fallback = [
    "#000000",
    "#FFFFFF",
    "#1F1F1F",
    "#EEECE1",
    "#4472C4",
    "#ED7D31",
    "#A5A5A5",
    "#FFC000",
    "#5B9BD5",
    "#70AD47",
  ];

  try {
    const bytes = await getDocumentBytes();
    const zip = await JSZip.loadAsync(bytes);
    const { presentationRelsDoc } = await getPresentationDocuments(zip);

    const slideMasterPath =
      findRelationshipTargetByType(presentationRelsDoc, "ppt/presentation.xml", PptRelationship.slideMaster);
    if (!slideMasterPath) {
      return { source: "fallback", colors: fallback, rowSize: fallback.length };
    }

    const slideMasterRelsPath = `${dirname(slideMasterPath)}/_rels/${fileNameWithoutExtension(slideMasterPath)}.xml.rels`;
    const slideMasterRelsXml = await readZipText(zip, slideMasterRelsPath);
    if (!slideMasterRelsXml) {
      return { source: "fallback", colors: fallback, rowSize: fallback.length };
    }

    const slideMasterRelsDoc = parseXml(slideMasterRelsXml);
    const themePath = findRelationshipTargetByType(slideMasterRelsDoc, slideMasterPath, PptRelationship.theme);
    if (!themePath) {
      return { source: "fallback", colors: fallback, rowSize: fallback.length };
    }

    const themeXml = await readZipText(zip, themePath);
    if (!themeXml) {
      return { source: "fallback", colors: fallback, rowSize: fallback.length };
    }

    const themeDoc = parseXml(themeXml);
    const schemeColors = extractThemeSchemeColors(themeDoc);
    if (schemeColors.length < 6) {
      return { source: "fallback", colors: fallback, rowSize: fallback.length };
    }

    const palette = flattenThemeRows(schemeColors.map((color) => deriveThemeScale(color)))
      .filter((color, index, array) => array.indexOf(color) === index);

    return {
      source: "deck",
      colors: palette,
      rowSize: schemeColors.length,
    };
  } catch {
    return {
      source: "fallback",
      colors: fallback,
      rowSize: fallback.length,
    };
  }
}

async function getCommentAuthors(zip: JSZip): Promise<Map<string, string>> {
  const authorsXml = await readZipText(zip, "ppt/commentAuthors.xml");
  const authors = new Map<string, string>();
  if (!authorsXml) {
    return authors;
  }

  const doc = parseXml(authorsXml);
  Array.from(doc.getElementsByTagName("*")).forEach((node) => {
    const name = localNameOf(node);
    if (name !== "cmAuthor" && name !== "commentAuthor") {
      return;
    }

    const id = node.getAttribute("id");
    const authorName = node.getAttribute("name") ?? node.getAttribute("initials") ?? "Unknown";
    if (id) {
      authors.set(id, authorName);
    }
  });

  return authors;
}

async function extractCommentsMarkdownFromZip(zip: JSZip): Promise<{ markdown: string; count: number }> {
  const structure = await getPresentationStructure(zip);
  const authors = await getCommentAuthors(zip);
  const sections: string[] = [];
  let count = 0;

  for (const slide of structure.slides) {
    if (!slide.commentsTarget) {
      continue;
    }

    const commentsXml = await readZipText(zip, slide.commentsTarget);
    if (!commentsXml) {
      continue;
    }

    const doc = parseXml(commentsXml);
    const entries = Array.from(doc.getElementsByTagName("*")).filter(
      (node): node is Element => {
        const name = localNameOf(node);
        return name === "cm" || name === "comment";
      },
    );

    const lines = entries
      .map((entry) => {
        const authorId = entry.getAttribute("authorId") ?? entry.getAttribute("author");
        const author = authorId ? authors.get(authorId) ?? authorId : "Unknown";
        const text = collectText(entry).join(" ").trim();
        if (!text) {
          return "";
        }
        count += 1;
        return `- ${author}: ${text}`;
      })
      .filter(Boolean);

    if (lines.length > 0) {
      sections.push(buildMarkdownSection(`${slide.index}. ${slide.title}`, lines.join("\n")));
    }
  }

  return {
    markdown: sections.length > 0 ? `# Presentation Comments\n\n${sections.join("\n")}` : "",
    count,
  };
}

async function extractNotesMarkdownFromZip(zip: JSZip): Promise<{ markdown: string; count: number }> {
  const structure = await getPresentationStructure(zip);
  const sections: string[] = [];
  let count = 0;

  for (const slide of structure.slides) {
    if (!slide.notesTarget) {
      continue;
    }

    const notesXml = await readZipText(zip, slide.notesTarget);
    if (!notesXml) {
      continue;
    }

    const doc = parseXml(notesXml);
    const rawLines = Array.from(doc.getElementsByTagName("*"))
      .filter((node): node is Element => localNameOf(node) === "t")
      .map((node) => node.textContent?.trim() ?? "")
      .filter(Boolean);

    const text = rawLines.join("\n").trim();
    if (!text) {
      continue;
    }

    count += 1;
    sections.push(buildMarkdownSection(`${slide.index}. ${slide.title}`, text));
  }

  return {
    markdown: sections.length > 0 ? `# Speaker Notes\n\n${sections.join("\n")}` : "",
    count,
  };
}

function removeRelationshipsBySuffix(relsDoc: XMLDocument, suffixes: string[]): string[] {
  const removedTargets: string[] = [];
  Array.from(relsDoc.getElementsByTagName("*")).forEach((node) => {
    if (localNameOf(node) !== "Relationship") {
      return;
    }

    const type = node.getAttribute("Type") ?? "";
    if (!suffixes.some((suffix) => type.endsWith(suffix))) {
      return;
    }

    const target = node.getAttribute("Target");
    if (target) {
      removedTargets.push(target);
    }
    node.parentNode?.removeChild(node);
  });

  return removedTargets;
}

function removeContentTypeOverrides(contentTypesDoc: XMLDocument, partNames: string[]) {
  const normalized = new Set(partNames.map((name) => `/${toPosixPath(name)}`));
  Array.from(contentTypesDoc.getElementsByTagName("*")).forEach((node) => {
    if (localNameOf(node) !== "Override") {
      return;
    }

    const partName = node.getAttribute("PartName");
    if (partName && normalized.has(partName)) {
      node.parentNode?.removeChild(node);
    }
  });
}

async function buildCleanedPresentationBase64(
  bytes: Uint8Array,
  mode: "comments" | "notes",
): Promise<{ base64: string; removedCount: number }> {
  const zip = await JSZip.loadAsync(bytes);
  const contentTypesXml = await readZipText(zip, "[Content_Types].xml");
  if (!contentTypesXml) {
    throw new Error("The presentation package is missing [Content_Types].xml.");
  }
  const contentTypesDoc = parseXml(contentTypesXml);
  const structure = await getPresentationStructure(zip);
  const removedParts = new Set<string>();
  let removedCount = 0;

  for (const slide of structure.slides) {
    const slideRelsPath = `${dirname(slide.target)}/_rels/${fileNameWithoutExtension(slide.target)}.xml.rels`;
    const slideRelsXml = await readZipText(zip, slideRelsPath);
    if (!slideRelsXml) {
      continue;
    }

    const slideRelsDoc = parseXml(slideRelsXml);
    const removedTargets = removeRelationshipsBySuffix(
      slideRelsDoc,
      mode === "comments" ? [PptRelationship.comments] : [PptRelationship.notesSlide],
    );

    if (removedTargets.length === 0) {
      continue;
    }

    removedCount += removedTargets.length;
    removedTargets.forEach((target) => {
      const resolvedTarget = resolvePartPath(slide.target, target);
      removedParts.add(resolvedTarget);
      if (resolvedTarget.includes("/notesSlides/")) {
        removedParts.add(`${dirname(resolvedTarget)}/_rels/${fileNameWithoutExtension(resolvedTarget)}.xml.rels`);
      }
    });

    zip.file(slideRelsPath, serializeXml(slideRelsDoc));
  }

  if (mode === "comments") {
    const presentationRelsPath = "ppt/_rels/presentation.xml.rels";
    const presentationRelsXml = await readZipText(zip, presentationRelsPath);
    if (presentationRelsXml) {
      const relsDoc = parseXml(presentationRelsXml);
      const removedTargets = removeRelationshipsBySuffix(relsDoc, [PptRelationship.commentAuthors]);
      removedTargets.forEach((target) => removedParts.add(resolvePartPath("ppt/presentation.xml", target)));
      zip.file(presentationRelsPath, serializeXml(relsDoc));
    }
  }

  removeContentTypeOverrides(contentTypesDoc, Array.from(removedParts));
  zip.file("[Content_Types].xml", serializeXml(contentTypesDoc));

  removedParts.forEach((part) => zip.remove(part));
  const base64 = await zip.generateAsync({ type: "base64" });
  return { base64, removedCount };
}

async function getSelectedSlideIndexes(): Promise<{ indexes: number[]; baseName: string }> {
  return PowerPoint.run(async (context) => {
    const slides = context.presentation.slides;
    const selectedSlides = context.presentation.getSelectedSlides();
    context.presentation.load("title");
    slides.load("items/id");
    selectedSlides.load("items/id");
    await context.sync();

    const idToIndex = new Map<string, number>();
    slides.items.forEach((slide, index) => idToIndex.set(slide.id, index + 1));
    const indexes = selectedSlides.items.map((slide) => idToIndex.get(slide.id)).filter((n): n is number => !!n);

    return {
      indexes,
      baseName: context.presentation.title || "Selected Slides",
    };
  });
}

async function buildSelectedSlidesPresentationBase64(
  bytes: Uint8Array,
  selectedIndexes: number[],
): Promise<{ base64: string; ignoredBrokenSlides: number }> {
  const zip = await JSZip.loadAsync(bytes);
  const { presentationDoc, presentationRelsDoc } = await getPresentationDocuments(zip);
  const selected = new Set(selectedIndexes);
  const slideRelationshipIds = new Set<string>();

  Array.from(presentationRelsDoc.getElementsByTagName("*")).forEach((node) => {
    if (localNameOf(node) !== "Relationship") {
      return;
    }

    const type = node.getAttribute("Type") ?? "";
    const id = node.getAttribute("Id");
    if (type.endsWith(PptRelationship.slide) && id) {
      slideRelationshipIds.add(id);
    }
  });

  const presentationRoot = presentationDoc.documentElement;
  childElements(presentationRoot).forEach((child) => {
    if (localNameOf(child) === "sectionLst" || localNameOf(child) === "custShowLst") {
      child.parentNode?.removeChild(child);
    }
  });

  const slideIdList = Array.from(presentationDoc.getElementsByTagName("*")).find((node): node is Element => {
    return localNameOf(node) === "sldIdLst";
  });

  if (!slideIdList) {
    throw new Error("The presentation package is missing its slide list.");
  }

  const keptRelIds = new Set<string>();
  let ignoredBrokenSlides = 0;
  childElements(slideIdList, "sldId").forEach((slideNode, index) => {
    const slideIndex = index + 1;
    const relId = getRelationshipId(slideNode);

    if (!relId || !slideRelationshipIds.has(relId)) {
      slideNode.parentNode?.removeChild(slideNode);
      ignoredBrokenSlides += 1;
      return;
    }

    if (!selected.has(slideIndex)) {
      slideNode.parentNode?.removeChild(slideNode);
      return;
    }
    keptRelIds.add(relId);
  });

  if (keptRelIds.size === 0) {
    throw new Error("Select at least one slide before creating a new deck.");
  }

  Array.from(presentationRelsDoc.getElementsByTagName("*")).forEach((node) => {
    if (localNameOf(node) !== "Relationship") {
      return;
    }

    const type = node.getAttribute("Type") ?? "";
    const id = node.getAttribute("Id") ?? "";
    if (type.endsWith(PptRelationship.slide) && !keptRelIds.has(id)) {
      node.parentNode?.removeChild(node);
    }
  });

  zip.file("ppt/presentation.xml", serializeXml(presentationDoc));
  zip.file("ppt/_rels/presentation.xml.rels", serializeXml(presentationRelsDoc));
  return {
    base64: await zip.generateAsync({ type: "base64" }),
    ignoredBrokenSlides,
  };
}

async function maybeCopyToClipboard(text: string): Promise<void> {
  if (!navigator.clipboard?.writeText) {
    throw new Error("Clipboard access is not available in this PowerPoint runtime.");
  }
  await navigator.clipboard.writeText(text);
}

async function isLocalBridgeAvailable(): Promise<boolean> {
  try {
    const response = await fetch(`${LOCAL_BRIDGE_ORIGIN}/healthz`, { method: "GET" });
    return response.ok;
  } catch {
    return false;
  }
}

async function postLocalBridge<T>(path: string, body: Record<string, unknown>): Promise<T> {
  const response = await fetch(`${LOCAL_BRIDGE_ORIGIN}${path}`, {
    method: "POST",
    headers: {
      "Content-Type": "application/json",
    },
    body: JSON.stringify(body),
  });

  const payload = await response.json();
  if (!response.ok) {
    throw new Error(payload?.error ?? "The local Jolify bridge request failed.");
  }

  return payload as T;
}

async function exportMarkdown(kind: "comments" | "notes"): Promise<{ markdown: string; count: number; baseName: string }> {
  const [bytes, fileProps] = await Promise.all([getDocumentBytes(), getFileProperties()]);
  const zip = await JSZip.loadAsync(bytes);
  const result = kind === "comments" ? await extractCommentsMarkdownFromZip(zip) : await extractNotesMarkdownFromZip(zip);
  return {
    markdown: result.markdown,
    count: result.count,
    baseName: getBaseNameFromUrl(fileProps.url),
  };
}

async function cleanDeck(kind: "comments" | "notes"): Promise<{ base64: string; removedCount: number; baseName: string }> {
  const [bytes, fileProps] = await Promise.all([getDocumentBytes(), getFileProperties()]);
  const cleaned = await buildCleanedPresentationBase64(bytes, kind);
  return {
    ...cleaned,
    baseName: getBaseNameFromUrl(fileProps.url),
  };
}

export async function exportPresentationAsPptx(): Promise<ActionResult> {
  const [bytes, fileProps] = await Promise.all([getDocumentBytes(), getFileProperties()]);
  const baseName = getBaseNameFromUrl(fileProps.url);
  const filename = `${sanitizeFilename(baseName)}.pptx`;

  if (await isLocalBridgeAvailable()) {
    await postLocalBridge<{ savedPath: string }>("/native/save-file", {
      base64File: toBase64(bytes),
      suggestedFilename: filename,
      openInPowerPoint: false,
    });

    return {
      type: "success",
      message: "Saved a PPTX copy of the current presentation through the local Jolify bridge.",
    };
  }

  downloadFile(filename, "application/vnd.openxmlformats-officedocument.presentationml.presentation", bytes);
  return {
    type: "success",
    message: "Downloaded a PPTX copy of the current presentation.",
  };
}

export async function exportPresentationAsPdf(): Promise<ActionResult> {
  const [bytes, fileProps] = await Promise.all([getDocumentBytes(Office.FileType.Pdf), getFileProperties()]);
  const baseName = getBaseNameFromUrl(fileProps.url);
  const filename = `${sanitizeFilename(baseName)}.pdf`;

  if (await isLocalBridgeAvailable()) {
    await postLocalBridge<{ savedPath: string }>("/native/save-file", {
      base64File: toBase64(bytes),
      suggestedFilename: filename,
      openInPowerPoint: false,
    });

    return {
      type: "success",
      message: "Saved a PDF copy of the current presentation through the local Jolify bridge.",
    };
  }

  downloadFile(filename, "application/pdf", bytes);
  return {
    type: "success",
    message: "Downloaded a PDF copy of the current presentation.",
  };
}

export async function copyCommentsToClipboard(): Promise<ActionResult> {
  const { markdown, count } = await exportMarkdown("comments");
  if (!markdown || count === 0) {
    return { type: "info", message: "No slide comments were found in this presentation." };
  }

  await maybeCopyToClipboard(markdown);
  return { type: "success", message: `Copied ${count} comment(s) to the clipboard.` };
}

export async function downloadCommentsMarkdown(): Promise<ActionResult> {
  const { markdown, count, baseName } = await exportMarkdown("comments");
  if (!markdown || count === 0) {
    return { type: "info", message: "No slide comments were found in this presentation." };
  }

  downloadFile(`${sanitizeFilename(baseName)} - comments.md`, "text/markdown;charset=utf-8", markdown);
  return { type: "success", message: `Downloaded comments for ${count} slide comment(s).` };
}

export async function cleanCommentsDeck(): Promise<ActionResult> {
  const { base64, removedCount, baseName } = await cleanDeck("comments");
  if (removedCount === 0) {
    return { type: "info", message: "No slide comments were found, so there was nothing to remove." };
  }

  const filename = `${sanitizeFilename(baseName)} - no comments.pptx`;
  if (await isLocalBridgeAvailable()) {
    await postLocalBridge<{ savedPath: string }>("/native/save-file", {
      base64File: base64,
      suggestedFilename: filename,
      openInPowerPoint: true,
    });

    return {
      type: "success",
      message: `Saved and opened a cleaned copy with ${removedCount} comment part(s) removed.`,
    };
  }

  downloadFile(filename, "application/vnd.openxmlformats-officedocument.presentationml.presentation", Uint8Array.from(atob(base64), (c) => c.charCodeAt(0)));
  return {
    type: "warning",
    message: `Downloaded a cleaned copy with ${removedCount} comment part(s) removed. Local mode is required to save and reopen it automatically.`,
  };
}

export async function copyNotesToClipboard(): Promise<ActionResult> {
  const { markdown, count } = await exportMarkdown("notes");
  if (!markdown || count === 0) {
    return { type: "info", message: "No speaker notes were found in this presentation." };
  }

  await maybeCopyToClipboard(markdown);
  return { type: "success", message: `Copied notes from ${count} slide(s) to the clipboard.` };
}

export async function downloadNotesMarkdown(): Promise<ActionResult> {
  const { markdown, count, baseName } = await exportMarkdown("notes");
  if (!markdown || count === 0) {
    return { type: "info", message: "No speaker notes were found in this presentation." };
  }

  downloadFile(`${sanitizeFilename(baseName)} - speaker notes.md`, "text/markdown;charset=utf-8", markdown);
  return { type: "success", message: `Downloaded notes for ${count} slide(s).` };
}

export async function cleanNotesDeck(): Promise<ActionResult> {
  const { base64, removedCount, baseName } = await cleanDeck("notes");
  if (removedCount === 0) {
    return { type: "info", message: "No speaker notes were found, so there was nothing to remove." };
  }

  const filename = `${sanitizeFilename(baseName)} - no notes.pptx`;
  if (await isLocalBridgeAvailable()) {
    await postLocalBridge<{ savedPath: string }>("/native/save-file", {
      base64File: base64,
      suggestedFilename: filename,
      openInPowerPoint: true,
    });

    return {
      type: "success",
      message: `Saved and opened a cleaned copy with ${removedCount} notes part(s) removed.`,
    };
  }

  downloadFile(filename, "application/vnd.openxmlformats-officedocument.presentationml.presentation", Uint8Array.from(atob(base64), (c) => c.charCodeAt(0)));
  return {
    type: "warning",
    message: `Downloaded a cleaned copy with ${removedCount} notes part(s) removed. Local mode is required to save and reopen it automatically.`,
  };
}

async function exportSelectedSlides(mode: SlideExportMode): Promise<ActionResult> {
  const [bytes, selection] = await Promise.all([getDocumentBytes(), getSelectedSlideIndexes()]);
  if (selection.indexes.length === 0) {
    return { type: "warning", message: "Select one or more slides before exporting them." };
  }

  const { base64, ignoredBrokenSlides } = await buildSelectedSlidesPresentationBase64(bytes, selection.indexes);
  const filename = `${sanitizeFilename(selection.baseName)} - ${selection.indexes.length} slides.pptx`;
  const ignoredBrokenSlidesNote =
    ignoredBrokenSlides > 0
      ? ` Ignored ${ignoredBrokenSlides} broken slide reference${ignoredBrokenSlides !== 1 ? "s" : ""} in the source deck.`
      : "";

  if (mode === "local-save" && (await isLocalBridgeAvailable())) {
    await postLocalBridge<{ savedPath: string }>("/native/save-file", {
      base64File: base64,
      suggestedFilename: filename,
      openInPowerPoint: false,
    });

    return {
      type: "success",
      message: `Saved a ${selection.indexes.length}-slide deck through the local Jolify bridge.${ignoredBrokenSlidesNote}`,
    };
  }

  downloadFile(filename, "application/vnd.openxmlformats-officedocument.presentationml.presentation", Uint8Array.from(atob(base64), (c) => c.charCodeAt(0)));
  return {
    type: mode === "local-save" ? "warning" : "success",
    message:
      mode === "local-save"
        ? `Downloaded a ${selection.indexes.length}-slide deck. Local mode is required for a native Save dialog.${ignoredBrokenSlidesNote}`
        : `Downloaded a ${selection.indexes.length}-slide deck.${ignoredBrokenSlidesNote}`,
  };
}

export function createDeckFromSelectedSlides(): Promise<ActionResult> {
  return exportSelectedSlides("local-save");
}

export async function createPresentationFromPictures(files: File[]): Promise<ActionResult> {
  if (files.length === 0) {
    return {
      type: "warning",
      message: "Choose one or more image files first.",
    };
  }

  if (!(await isLocalBridgeAvailable())) {
    return {
      type: "warning",
      message: "Create Presentation from Pictures currently needs local mode.",
    };
  }

  const supportedFiles = files.filter((file) => /^image\/(png|jpe?g|svg\+xml)$/i.test(file.type));
  if (supportedFiles.length === 0) {
    return {
      type: "warning",
      message: "Choose PNG, JPG, JPEG, or SVG image files.",
    };
  }

  const images = await Promise.all(
    supportedFiles.map(async (file) => ({
      filename: file.name,
      base64Image: await fileToBase64(file),
    })),
  );

  const suggestedFilename = `${sanitizeFilename(supportedFiles[0].name.replace(/\.[^.]+$/, ""))} - pictures.pptx`;
  await postLocalBridge<{ savedPath: string }>("/native/create-presentation-from-pictures", {
    images,
    suggestedFilename,
  });

  return {
    type: "success",
    message: `Created and opened a new ${supportedFiles.length}-slide presentation from the selected picture${supportedFiles.length !== 1 ? "s" : ""}.`,
  };
}

export async function attachSelectedSlidesToEmail(): Promise<ActionResult> {
  const [bytes, selection] = await Promise.all([getDocumentBytes(), getSelectedSlideIndexes()]);
  if (selection.indexes.length === 0) {
    return { type: "warning", message: "Select one or more slides before attaching them to an email." };
  }

  if (!(await isLocalBridgeAvailable())) {
    return {
      type: "warning",
      message: "This workflow needs local mode. Use Create Deck from Selected Slides for a download in hosted mode.",
    };
  }

  const { base64 } = await buildSelectedSlidesPresentationBase64(bytes, selection.indexes);
  const filename = `${sanitizeFilename(selection.baseName)} - ${selection.indexes.length} slides.pptx`;
  const saveResult = await postLocalBridge<{ savedPath: string }>("/native/save-file", {
    base64File: base64,
    suggestedFilename: filename,
    openInPowerPoint: false,
  });

  await postLocalBridge<{ ok: boolean }>("/native/create-outlook-draft", {
    attachmentPath: saveResult.savedPath,
    subject: `${selection.baseName} - ${selection.indexes.length} slides`,
  });

  return {
    type: "success",
    message: `Saved the slide deck and opened an Outlook draft with it attached.`,
  };
}
