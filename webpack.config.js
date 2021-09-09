const devCerts = require("office-addin-dev-certs");
const CopyWebpackPlugin = require("copy-webpack-plugin");
const HtmlWebpackPlugin = require("html-webpack-plugin");
require("dotenv").config();

const urlPlaceholder = "{blobStore}";

module.exports = async (env, options) => {
  return {
    devtool: "source-map",
    entry: {
      commands: "./src/commands.ts",
    },
    resolve: {
      extensions: [".ts", ".tsx", ".html", ".js"],
    },
    module: {
      rules: [
        {
          test: /\.ts$/,
          exclude: /node_modules/,
          use: "babel-loader",
        },
        {
          test: /\.s[ac]ss$/i,
          use: [
            // Creates `style` nodes from JS strings
            "style-loader",
            // Translates CSS into CommonJS
            "css-loader",
            // Compiles Sass to CSS
            "sass-loader",
          ],
        },
        {
          test: /\.tsx?$/,
          exclude: /node_modules/,
          use: "ts-loader",
        },
        {
          test: /\.html$/,
          exclude: /node_modules/,
          use: "html-loader",
        },
        {
          test: /\.(png|jpg|jpeg|gif)$/,
          loader: "file-loader",
          options: {
            name: "[path][name].[ext]",
          },
        },
      ],
    },
    plugins: [
      new HtmlWebpackPlugin({
        filename: "commands.html",
        template: "./src/commands.html",
        chunks: ["commands"],
      }),
      new CopyWebpackPlugin({
        patterns: [
          {
            to: "[name][ext]",
            from: "src/manifest.xml",
            transform(content) {
              return content
                .toString()
                .replace(
                  new RegExp(urlPlaceholder, "g"),
                  process.env.URL || "localhost:3000"
                );
            },
          },
        ],
      }),
    ],
    devServer: {
      headers: {
        "Access-Control-Allow-Origin": "*",
      },
      https:
        options.https !== undefined
          ? options.https
          : await devCerts.getHttpsServerOptions().then((config) => {
              // Unsuported key.
              delete config.ca;
              return config;
            }),
      port: process.env.npm_package_config_dev_server_port || 3000,
    },
  };
};
