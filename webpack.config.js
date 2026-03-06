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
            from: "install.sh",
            to: "install.sh",
          },
          {
            from: "uninstall.sh",
            to: "uninstall.sh",
          },
          {
            from: "manifest*.xml",
            to: "[name]" + "[ext]",
            transform(content) {
              if (dev) {
                return content;
              } else {
                return content.toString().replace(new RegExp(urlDev, "g"), urlProd);
              }
            },
          },
        ],
      }),
      new HtmlWebpackPlugin({
        filename: "commands.html",
        template: "./src/commands/commands.html",
        chunks: ["polyfill", "commands"],
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
        filename: "dialogs/rename-shapes.html",
        template: "./src/dialogs/rename-shapes.html",
        chunks: [],
        inject: false,
      }),
      new HtmlWebpackPlugin({
        filename: "dialogs/slide-outline.html",
        template: "./src/dialogs/slide-outline.html",
        chunks: [],
        inject: false,
      }),
      new HtmlWebpackPlugin({
        filename: "dialogs/gantt-builder.html",
        template: "./src/dialogs/gantt-builder.html",
        chunks: [],
        inject: false,
      }),
      new HtmlWebpackPlugin({
        filename: "dialogs/timeline-builder.html",
        template: "./src/dialogs/timeline-builder.html",
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
      client: {
        // tell WDS client where to open the websocket
        webSocketURL: {
          protocol: "wss",
          hostname: "localhost",
          port: 3300,
          pathname: "/ws",
        },
        overlay: true,
        logging: "info",
      },
      allowedHosts: "all",
      // If you still see WS issues, you can disable hot reload:
      // hot: false,
      // liveReload: false,
    },
  };

  return config;
};
