const path = require("path");
const CopyWebpackPlugin = require("copy-webpack-plugin");
const HtmlWebpackPlugin = require("html-webpack-plugin");

const urlDev = "http://localhost:3001/";

const addinConfig = (env, options) => {
  const dev = options.mode === "development";
  return {
    devtool: "source-map",
    entry: {
      polyfill: ["core-js/stable", "regenerator-runtime/runtime"],
      taskpane: "./src/taskpane/taskpane.ts",
      commands: "./src/commands/commands.ts",
    },
    output: {
      devtoolModuleFilenameTemplate: "webpack:///[resource-path]?[loaders]",
      path: path.resolve(__dirname, "dist"),
    },
    resolve: {
      extensions: [".ts", ".tsx", ".html", ".js"],
    },
    module: {
      rules: [
        {
          test: /\.ts$/,
          exclude: /node_modules/,
          use: {
            loader: "ts-loader",
            options: {
              transpileOnly: true,
            },
          },
        },
        {
          test: /\.tsx?$/,
          exclude: /node_modules/,
          use: "ts-loader",
        },
        {
          test: /\.css$/i,
          use: ["style-loader", "css-loader"],
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
      new HtmlWebpackPlugin({
        filename: "taskpane.html",
        template: "./src/taskpane/taskpane.html",
        chunks: ["polyfill", "taskpane"],
      }),
      new HtmlWebpackPlugin({
        filename: "commands.html",
        template: "./src/commands/commands.html",
        chunks: ["polyfill", "commands"],
      }),
      new CopyWebpackPlugin({
        patterns: [
          {
            from: "manifest.xml",
            to: dev ? "manifest.dev.xml" : "manifest.xml",
            transform(content) {
              if (dev) {
                // Replace production URLs with localhost for development
                return content
                  .toString()
                  .replace(/https:\/\/ipg-taxonomy-extractor-addin\.wookstar\.com\/taskpane/g, urlDev + 'taskpane.html')
                  .replace(/https:\/\/ipg-taxonomy-extractor-addin\.wookstar\.com\/commands/g, urlDev + 'commands.html')
                  .replace(/https:\/\/ipg-taxonomy-extractor-addin\.wookstar\.com\/assets/g, urlDev + 'assets')
                  .replace(/https:\/\/ipg-taxonomy-extractor-addin\.wookstar\.com/g, urlDev.slice(0, -1));
              }
              return content;
            },
          },
          {
            from: "assets",
            to: "assets",
          },
        ],
      }),
    ],
    devServer: {
      hot: true,
      headers: {
        "Access-Control-Allow-Origin": "*",
      },
      server: "http",
      port: 3001,
      historyApiFallback: {
        index: "/taskpane.html",
      },
      static: {
        directory: "./dist",
        publicPath: "/",
      },
      open: false,
    },
  };
};

module.exports = addinConfig;