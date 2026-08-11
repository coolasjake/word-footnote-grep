/* eslint-disable @typescript-eslint/no-var-requires */
const path = require("path");
const HtmlWebpackPlugin = require("html-webpack-plugin");
const CopyWebpackPlugin = require("copy-webpack-plugin");

const devCerts = (() => {
  try {
    return require("office-addin-dev-certs");
  } catch {
    return null;
  }
})();

module.exports = async (env, options) => {
  const isDev = options.mode === "development";

  let httpsOptions = {};
  if (isDev && devCerts) {
    httpsOptions = await devCerts.getHttpsServerOptions();
  }

  return {
    entry: {
      taskpane: "./src/taskpane/taskpane.ts",
    },
    output: {
      path: path.resolve(__dirname, "dist"),
      filename: "[name].js",
      clean: true,
    },
    resolve: {
      extensions: [".ts", ".js"],
    },
    module: {
      rules: [
        {
          test: /\.ts$/,
          use: "ts-loader",
          exclude: /node_modules/,
        },
        {
          test: /\.css$/,
          use: ["style-loader", "css-loader"],
        },
        {
          test: /\.html$/,
          loader: "html-loader",
        },
      ],
    },
    plugins: [
      new HtmlWebpackPlugin({
        template: "./src/taskpane/taskpane.html",
        filename: "taskpane.html",
        chunks: ["taskpane"],
      }),
      new CopyWebpackPlugin({
        patterns: [
          { from: "assets", to: "assets", noErrorOnMissing: true },
          { from: "src/commands.html", to: "commands.html" },
        ],
      }),
    ],
    devServer: {
      static: path.resolve(__dirname, "dist"),
      headers: {
        "Access-Control-Allow-Origin": "*",
      },
      server: isDev
        ? {
            type: "https",
            options: httpsOptions,
          }
        : undefined,
      port: 3000,
    },
    devtool: isDev ? "source-map" : false,
  };
};
