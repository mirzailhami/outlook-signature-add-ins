const path = require("path");
const CopyWebpackPlugin = require("copy-webpack-plugin");
const HtmlWebpackPlugin = require("html-webpack-plugin");
const devCerts = require("office-addin-dev-certs");
const webpack = require("webpack");

const getHttpsOptions = async () => {
  try {
    const httpsOptions = await devCerts.getHttpsServerOptions();
    return { ca: httpsOptions.ca, key: httpsOptions.key, cert: httpsOptions.cert };
  } catch (error) {
    console.error("Failed to get HTTPS certificates, falling back to HTTP:", error.message);
    return {};
  }
};

// Pre-fetch HTTPS options to avoid async issues
const httpsOptions = getHttpsOptions();

module.exports = (env, options) => {
  const isProduction = options.mode === "production";
  const envFile = isProduction ? ".env.production" : ".env";
  const envPath = path.resolve(__dirname, envFile);
  const envVars = require("dotenv").config({ path: envPath }).parsed || {};

  const config = {
    devtool: "source-map",
    entry: {
      polyfill: ["core-js/stable", "regenerator-runtime/runtime"],
      commands: "./src/commands/commands.js",
    },
    output: {
      clean: true,
      path: path.resolve(__dirname, "dist"),
      filename: "[name].[contenthash].js",
    },
    resolve: {
      extensions: [".js", ".html"],
    },
    module: {
      rules: [
        {
          test: /\.js$/,
          exclude: /node_modules/,
          use: {
            loader: "babel-loader",
          },
        },
        {
          test: /\.html$/,
          exclude: /node_modules/,
          use: {
            loader: "html-loader",
          },
        },
      ],
    },
    plugins: [
      new webpack.DefinePlugin({
        "process.env": JSON.stringify(envVars),
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
            to: "manifest.xml",
            transform(content) {
              const assetBaseUrl = process.env.ASSET_BASE_URL || "https://localhost:3000";
              if (!assetBaseUrl) {
                throw new Error("ASSET_BASE_URL is not defined");
              }
              return content.toString().replace(/\${ASSET_BASE_URL}/g, assetBaseUrl);
            },
          },
          {
            from: "src/index.html",
            to: "index.html",
          },
          {
            from: "src/taskpane/taskpane.html",
            to: "taskpane.html",
          },
          {
            from: "src/commands/commands.js",
            to: "commands.js",
          },
        ],
      }),
    ],
    devServer: {
      hot: true,
      headers: {
        "Access-Control-Allow-Origin": "*",
      },
      server: {
        type: "https",
        options: httpsOptions,
      },
      port: 3000,
      devMiddleware: {
        writeToDisk: true,
      },
      allowedHosts: ["localhost", ".azurewebsites.net", ".ngrok-free.app"],
    },
    mode: options.mode || "development",
  };

  return config;
};
