const path = require("path");
const devCerts = require("office-addin-dev-certs");
const CopyWebpackPlugin = require("copy-webpack-plugin");
const HtmlWebpackPlugin = require("html-webpack-plugin");
const webpack = require("webpack");

async function getHttpsOptions() {
  const httpsOptions = await devCerts.getHttpsServerOptions();
  return { ca: httpsOptions.ca, key: httpsOptions.key, cert: httpsOptions.cert };
}

module.exports = async (env, options) => {
  const isProduction = options.mode === "production";
  const envFile = isProduction ? ".env.production" : ".env";
  const envPath = path.resolve(__dirname, envFile);
  const envVars = require("dotenv").config({ path: envPath }).parsed || {};

  const config = {
    devtool: "source-map",
    entry: {
      polyfill: ["core-js", "regenerator-runtime/runtime"],
      vendor: ["react", "react-dom", "@fluentui/react-components", "@fluentui/react-icons"],
      taskpane: ["./src/taskpane/index.jsx", "./src/taskpane/taskpane.html"],
      commands: "./src/commands/commands.js",
      error: "./src/error/error.js",
    },
    output: {
      clean: true,
      path: path.resolve(__dirname, "dist"),
      filename: "[name].[contenthash].js", // Use hashed filenames
    },
    resolve: {
      extensions: [".js", ".jsx", ".html"],
    },
    module: {
      rules: [
        {
          test: /\.jsx?$/,
          use: "babel-loader",
          exclude: /node_modules/,
        },
        {
          test: /\.html$/,
          exclude: /node_modules/,
          use: "html-loader",
        }
      ],
    },
    plugins: [
      new webpack.DefinePlugin({
        "process.env": JSON.stringify(envVars),
      }),
      new CopyWebpackPlugin({
        patterns: [
          {
            from: "manifest*.xml",
            to: "[name][ext]",
            transform(content) {
              console.log("Transforming manifest.xml");
              if (!process.env.ASSET_BASE_URL) {
                throw new Error(`ASSET_BASE_URL is not defined in ${envFile}`);
              }
              const result = content.toString().replace(/\${ASSET_BASE_URL}/g, process.env.ASSET_BASE_URL);
              console.log("Transformed manifest.xml sample:", result.slice(0, 200));
              return result;
            },
          },
        ],
      }),
      new HtmlWebpackPlugin({
        filename: "taskpane.html",
        template: "./src/taskpane/taskpane.html",
        chunks: ["polyfill", "vendor", "taskpane"],
      }),
      new HtmlWebpackPlugin({
        filename: "commands.html",
        template: "./src/commands/commands.html",
        chunks: ["polyfill", "commands"],
      }),
      new HtmlWebpackPlugin({
        filename: "error.html",
        template: "./src/error/error.html",
        chunks: ["polyfill", "vendor", "error"],
      }),
    ],
    devServer: {
      hot: true,
      headers: {
        "Access-Control-Allow-Origin": "*",
        "Access-Control-Allow-Methods": "GET, POST, OPTIONS",
        "Access-Control-Allow-Headers": "Content-Type",
      },
      server: {
        type: "https",
        options: env.WEBPACK_BUILD || options.https !== undefined ? options.https : await getHttpsOptions(),
      },
      port: process.env.npm_package_config_dev_server_port || 3000,
      devMiddleware: {
        writeToDisk: true,
      },
      allowedHosts: ["localhost", ".azurewebsites.net", ".ngrok-free.app"],
    },
  };

  console.log("ASSET_BASE_URL after Dotenv:", process.env.ASSET_BASE_URL);
  return config;
};