const devCerts = require("office-addin-dev-certs");
const CopyWebpackPlugin = require("copy-webpack-plugin");
const HtmlWebpackPlugin = require("html-webpack-plugin");
require("dotenv").config();

module.exports = async (env, options) => {
  return {
    devtool: "source-map",
    entry: {
      commands: "./src/commands/commands.ts",
    },
    resolve: {
      extensions: [".ts"],
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
    plugins: [
      new HtmlWebpackPlugin({
        filename: "commands.html",
        template: "./src/commands/commands.html",
        chunks: ["commands"],
      }),
      new CopyWebpackPlugin({
        patterns: [
          {
            to: "[name][ext]",
            from: "manifest.xml",
            transform(content) {
              return content
                .toString()
                .replace(new RegExp("{blobStore}", "g"), process.env.URL)
                .replace(new RegExp("{apiUrl}", "g"), process.env.API)
                .replace(new RegExp("{appId}", "g"), process.env.APPID);
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
