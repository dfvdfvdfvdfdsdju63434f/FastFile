const path = require("path");
const CopyWebpackPlugin = require("copy-webpack-plugin");

module.exports = (env, argv) => {
  const isProduction = argv.mode === "production";

  return {
    devtool: isProduction ? false : "source-map",
    entry: {
      taskpane: "./src/taskpane/taskpane.ts",
      commands: "./src/commands/commands.ts",
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
      ],
    },
    plugins: [
      new CopyWebpackPlugin({
        patterns: [
          { from: "src/taskpane/taskpane.html", to: "taskpane.html" },
          { from: "src/taskpane/taskpane.css", to: "taskpane.css" },
          { from: "src/commands/commands.html", to: "commands.html" },
          { from: "assets", to: "assets" },
          { from: "manifest.xml", to: "manifest.xml" },
        ],
      }),
    ],
    devServer: {
      static: { directory: path.join(__dirname, "dist") },
      server: "https", // uses office-addin-dev-certs automatically via office-addin-debugging
      port: 3000,
      hot: true,
      headers: {
        "Access-Control-Allow-Origin": "*",
      },
    },
  };
};
