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