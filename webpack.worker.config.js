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
    library: {
      type: "module",
    },
  },
  experiments: {
    outputModule: true,
  },
  resolve: {
    extensions: [".ts", ".js"],
  },
  module: {
    rules: [
      {
        test: /\.ts$/,
        exclude: /node_modules/,
        use: {
          loader: "ts-loader",
          options: {
            configFile: "tsconfig.worker.json",
            transpileOnly: false,
          },
        },
      },
    ],
  },
};