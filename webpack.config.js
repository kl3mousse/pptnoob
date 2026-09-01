const path = require("path");
const HtmlWebpackPlugin = require("html-webpack-plugin");
const CopyWebpackPlugin = require("copy-webpack-plugin");

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
      new CopyWebpackPlugin({
        patterns: [{ from: "assets", to: "assets" }],
      }),
    ],
    devServer: {
      server: { type: "https", options: https },
      port: 3000,
      headers: { "Access-Control-Allow-Origin": "*" },
    },
  };
};
