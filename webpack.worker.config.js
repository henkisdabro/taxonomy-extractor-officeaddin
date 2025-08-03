const path = require("path");

module.exports = {
  entry: {
    worker: "./src/worker.ts",
  },
  target: "webworker",
  mode: "production",
  output: {
    filename: "[name].js",
    path: path.resolve(__dirname, "dist"),
  },
  resolve: {
    extensions: [".ts", ".js"],
  },
  module: {
    rules: [
      {
        test: /\.ts$/,
        exclude: /node_modules/,
        use: "ts-loader",
      },
    ],
  },
};