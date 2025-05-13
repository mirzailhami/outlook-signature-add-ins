const path = require("path");
const webpack = require("webpack");
const { CleanWebpackPlugin } = require("clean-webpack-plugin");
const CopyWebpackPlugin = require("copy-webpack-plugin");
const HtmlWebpackPlugin = require("html-webpack-plugin");
const dotenv = require("dotenv");
const devCerts = require("office-addin-dev-certs");

const getHttpsOptions = async () => {
  try {
    const httpsOptions = await devCerts.getHttpsServerOptions();
    return { ca: httpsOptions.ca, key: httpsOptions.key, cert: httpsOptions.cert };
  } catch (error) {
    console.error("Failed to get HTTPS certificates, falling back to HTTP:", error.message);
    return {};
  }
};

module.exports = async (env, options) => {
  const isProduction = options.mode === "production";

  // Load environment variables
  const envPath = path.resolve(__dirname, isProduction ? ".env.production" : ".env");
  const envVars = dotenv.config({ path: envPath }).parsed || {};

  // Fallback to appropriate URL
  const assetBaseUrl = isProduction
    ? envVars.ASSET_BASE_URL || "https://mirzailhami.github.io/outlook-signature-add-ins"
    : envVars.ASSET_BASE_URL || "https://localhost:3000";

  // Get HTTPS options for dev server
  const httpsOptions = await getHttpsOptions();

  return {
    mode: isProduction ? "production" : "development",
    devtool: isProduction ? "source-map" : "eval-source-map",

    entry: {
      polyfill: ["core-js/stable", "regenerator-runtime/runtime"],
      commands: "./src/commands/commands.js",
      taskpane: "./src/taskpane/taskpane.js",
      launchevent: "./src/launchevent/launchevent.js",
    },

    output: {
      path: path.resolve(__dirname, "dist"),
      filename: "[name].[contenthash:8].js",
      publicPath: assetBaseUrl,
      clean: true,
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
            options: {
              presets: [
                [
                  "@babel/preset-env",
                  {
                    useBuiltIns: "usage",
                    corejs: 3,
                    targets: {
                      browsers: ["last 2 versions", "not dead", "not ie 11"],
                    },
                  },
                ],
              ],
            },
          },
        },
        {
          test: /\.html$/,
          use: ["html-loader"],
        },
      ],
    },

    plugins: [
      new CleanWebpackPlugin(),
      new webpack.DefinePlugin({
        "process.env.ASSET_BASE_URL": JSON.stringify(assetBaseUrl),
      }),
      // Taskpane HTML
      new HtmlWebpackPlugin({
        template: "./src/taskpane/taskpane.html",
        filename: "taskpane.html",
        chunks: ["polyfill", "taskpane"],
        publicPath: assetBaseUrl,
        minify: isProduction
          ? {
              removeComments: true,
              collapseWhitespace: true,
              removeRedundantAttributes: true,
            }
          : false,
      }),
      new HtmlWebpackPlugin({
        template: "./src/commands/commands.html",
        filename: "commands.html",
        chunks: ["polyfill", "commands", "launchevent"],
        publicPath: assetBaseUrl,
        minify: isProduction
          ? {
              removeComments: true,
              collapseWhitespace: true,
              removeRedundantAttributes: true,
            }
          : false,
      }),
      new CopyWebpackPlugin({
        patterns: [
          {
            from: "manifest.xml",
            transform(content) {
              return content.toString().replace(/\${ASSET_BASE_URL}/g, assetBaseUrl);
            },
          },
          {
            from: ".nojekyll",
            to: ".nojekyll",
          },
          {
            from: "src/index.html",
            to: "index.html",
          },
          {
            from: "assets",
            to: "assets",
          },
        ],
      }),
    ],
    devServer: {
      static: {
        directory: path.join(__dirname, "dist"),
      },
      compress: true,
      port: 3000,
      hot: true,
      headers: {
        "Access-Control-Allow-Origin": "*",
        "Access-Control-Allow-Methods": "GET, POST, PUT, DELETE, PATCH, OPTIONS",
        "Access-Control-Allow-Headers": "X-Requested-With, content-type, Authorization",
      },
      allowedHosts: ["localhost", ".azurewebsites.net"],
      server: {
        type: "https",
        options: httpsOptions,
      },
      devMiddleware: {
        writeToDisk: true,
      },
    },

    performance: {
      hints: isProduction ? "warning" : false,
      maxAssetSize: 1024 * 1024,
      maxEntrypointSize: 1024 * 1024,
    },
  };
};
