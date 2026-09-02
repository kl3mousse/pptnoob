const path = require("path");
const fs = require("fs");
const webpack = require("webpack");
const HtmlWebpackPlugin = require("html-webpack-plugin");
const CopyWebpackPlugin = require("copy-webpack-plugin");

const phosphorSvgDirectory = path.resolve(__dirname, "assets/phosphor/phosphor-icons/SVGs");
const phosphorWeights = ["thin", "light", "regular", "bold", "fill", "duotone"];

function createPhosphorIconIndex() {
  return Object.fromEntries(phosphorWeights.map((weight) => [
    weight,
    fs.readdirSync(path.join(phosphorSvgDirectory, weight))
      .filter((filename) => filename.endsWith(".svg"))
      .sort(),
  ]));
}

class PhosphorIconIndexPlugin {
  apply(compiler) {
    compiler.hooks.thisCompilation.tap("PhosphorIconIndexPlugin", (compilation) => {
      compilation.hooks.processAssets.tap(
        {
          name: "PhosphorIconIndexPlugin",
          stage: webpack.Compilation.PROCESS_ASSETS_STAGE_ADDITIONAL,
        },
        () => {
          const index = JSON.stringify(createPhosphorIconIndex());
          compilation.emitAsset("assets/phosphor/icon-index.json", new webpack.sources.RawSource(index));
        },
      );
    });
  }
}

module.exports = async (env, options) => {
  const dev = options.mode === "development";

  // Use locally-installed dev certs in development
  const https = dev
    ? await require("office-addin-dev-certs").getHttpsServerOptions()
    : undefined;

  return {
    entry: { taskpane: "./src/taskpane/taskpane.ts" },
    output: {
      path: path.resolve(__dirname, "dist"),
      filename: "[name].js",
      clean: true,
    },
    resolve: { extensions: [".ts", ".js"] },
    module: {
      rules: [{ test: /\.ts$/, use: "ts-loader", exclude: /node_modules/ }],
    },
    plugins: [
      new HtmlWebpackPlugin({
        filename: "taskpane.html",
        template: "./src/taskpane/taskpane.html",
        chunks: ["taskpane"],
      }),
      new HtmlWebpackPlugin({
        filename: "info.html",
        template: "./src/info/info.html",
        inject: false,
      }),
      new CopyWebpackPlugin({
        patterns: [{ from: "assets", to: "assets" }],
      }),
      new PhosphorIconIndexPlugin(),
    ],
    devServer: {
      server: { type: "https", options: https },
      port: 3000,
      headers: { "Access-Control-Allow-Origin": "*" },
    },
  };
};
