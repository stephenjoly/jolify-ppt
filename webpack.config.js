/* eslint-disable no-undef */

const devCerts = require("office-addin-dev-certs");
const CopyWebpackPlugin = require("copy-webpack-plugin");
const HtmlWebpackPlugin = require("html-webpack-plugin");

const urlDev = "https://localhost:3300/";
const urlProd = "https://stephenjoly.github.io/jolify-ppt/";

async function getHttpsOptions() {
  const httpsOptions = await devCerts.getHttpsServerOptions();
  return { ca: httpsOptions.ca, key: httpsOptions.key, cert: httpsOptions.cert };
}

module.exports = async (env, options) => {
  const dev = options.mode === "development";
  const config = {
    devtool: "source-map",
    entry: {
      polyfill: ["core-js/stable", "regenerator-runtime/runtime"],
      commands: "./src/commands/commands.ts",
      cleanupDeckDialog: "./src/dialogs/cleanup-deck-dialog.ts",
      selectedDeckDialog: "./src/dialogs/selected-deck-dialog.ts",
      weekdayRangeDialog: "./src/dialogs/weekday-range-dialog.ts",
      taskpane: "./src/taskpane/taskpane.ts",
    },
    output: {
      clean: true,
      filename: "[name].js", // => commands.js, polyfill.js
      path: require("path").resolve(__dirname, "dist"),
      publicPath: dev ? "/" : urlProd,
    },
    resolve: {
      extensions: [".ts", ".html", ".js"],
    },
    module: {
      rules: [
        {
          test: /\.ts$/,
          exclude: /node_modules/,
          use: {
            loader: "babel-loader",
          },
        },
        {
          test: /\.html$/,
          exclude: /node_modules/,
          use: "html-loader",
        },
        {
          test: /\.(png|jpg|jpeg|gif|ico)$/,
          type: "asset/resource",
          generator: {
            filename: "assets/[name][ext][query]",
          },
        },
      ],
    },
    plugins: [
      new CopyWebpackPlugin({
        patterns: [
          {
            from: "assets",
            to: "assets",
          },
          {
            from: "index.html",
            to: "index.html",
          },
          {
            from: "guide.html",
            to: "guide.html",
          },
          {
            from: "install/install.sh",
            to: "install.sh",
          },
          {
            from: "install/uninstall.sh",
            to: "uninstall.sh",
          },
          {
            from: "install/install-local.sh",
            to: "install-local.sh",
          },
          {
            from: "install/uninstall-local.sh",
            to: "uninstall-local.sh",
          },
          {
            from: dev ? "dev/manifest.xml" : "manifest.xml",
            to: dev ? "manifest.dev.xml" : "[name][ext]",
            transform(content) {
              if (dev) {
                return content;
              }
              return content.toString().replace(new RegExp(urlDev, "g"), urlProd);
            },
          },
        ],
      }),
      new HtmlWebpackPlugin({
        filename: "commands.html",
        chunks: ["polyfill", "commands"],
        inject: false,
        templateContent: () => `<!doctype html>
<html>
  <head>
    <meta charset="utf-8" />
    <script src="https://appsforoffice.microsoft.com/lib/1/hosted/office.js"></script>
    <script src="${dev ? urlDev : urlProd}polyfill.js"></script>
    <script src="${dev ? urlDev : urlProd}commands.js"></script>
  </head>
  <body></body>
</html>`,
      }),
      new HtmlWebpackPlugin({
        filename: "taskpane.html",
        template: "./src/taskpane/taskpane.html",
        chunks: ["polyfill", "taskpane"],
      }),
      new HtmlWebpackPlugin({
        filename: "dialogs/grid-builder.html",
        template: "./src/dialogs/grid-builder.html",
        chunks: [],
        inject: false,
      }),
      new HtmlWebpackPlugin({
        filename: "dialogs/cleanup-deck.html",
        template: "./src/dialogs/cleanup-deck-dialog.html",
        chunks: ["polyfill", "cleanupDeckDialog"],
      }),
      new HtmlWebpackPlugin({
        filename: "dialogs/selected-deck.html",
        template: "./src/dialogs/selected-deck-dialog.html",
        chunks: ["polyfill", "selectedDeckDialog"],
      }),
      new HtmlWebpackPlugin({
        filename: "dialogs/weekday-range.html",
        template: "./src/dialogs/weekday-range-dialog.html",
        chunks: ["polyfill", "weekdayRangeDialog"],
      }),
      new HtmlWebpackPlugin({
        filename: "dialogs/symbol-picker.html",
        template: "./src/dialogs/symbol-picker.html",
        chunks: [],
        inject: false,
      }),
    ],
    devServer: {
      headers: { "Access-Control-Allow-Origin": "*" },
      server: {
        type: "https",
        options:
          env.WEBPACK_BUILD || options.https !== undefined
            ? options.https
            : await getHttpsOptions(),
      },
      host: "localhost", // force localhost (not 0.0.0.0)
      port: process.env.npm_package_config_dev_server_port || 3300,
      static: [
        // serve emitted files and your assets folder
        { directory: __dirname + "/dist" },
        { directory: __dirname + "/assets" },
      ],
      client: false,
      allowedHosts: "all",
      hot: false,
      liveReload: false,
    },
  };

  return config;
};
