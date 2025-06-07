const path = require("path");
const webpack = require("webpack");
const { CleanWebpackPlugin } = require("clean-webpack-plugin");
const CopyWebpackPlugin = require("copy-webpack-plugin");
const HtmlWebpackPlugin = require("html-webpack-plugin");
const Dotenv = require("dotenv-webpack");
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

  // Get HTTPS options for dev server
  const httpsOptions = await getHttpsOptions();

  return {
    mode: isProduction ? "production" : "development",
    devtool: isProduction ? "source-map" : "eval-source-map",
    stats: { errorDetails: true }, // Enable detailed error messages

    entry: {
      polyfill: ["core-js/stable", "regenerator-runtime/runtime"],
      commands: ["./src/commands/commands.js", "./src/commands/commands.html"],
      taskpane: ["./src/taskpane/taskpane.html"],
    },

    output: {
      clean: true,
    },

    resolve: {
      extensions: [".js", ".html"],
    },

    module: {
      rules: [
        {
          test: /\.js$/,
          include: [
            path.resolve(__dirname, "node_modules/@azure"),
            path.resolve(__dirname, "node_modules/@microsoft"),
            path.resolve(__dirname, "src"),
          ],
          use: {
            loader: "babel-loader",
            options: {
              presets: [
                [
                  "@babel/preset-env",
                  {
                    targets: {
                      browsers: ["since 2015"],
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
          exclude: /node_modules/,
        },
      ],
    },

    plugins: [
      new CleanWebpackPlugin(),
      new Dotenv({
        path: path.resolve(__dirname, isProduction ? ".env.production" : ".env"),
        safe: true,
        allowEmptyValues: true,
        systemvars: true,
      }),
      new HtmlWebpackPlugin({
        template: "./src/taskpane/taskpane.html",
        filename: "taskpane.html",
        chunks: ["polyfill", "commands"],
        publicPath: isProduction ? "/outlook-signature-add-ins/" : "/",
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
        chunks: ["polyfill", "commands"],
        publicPath: isProduction ? "/outlook-signature-add-ins/" : "/",
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
              const assetBaseUrl =
                process.env.ASSET_BASE_URL ||
                (isProduction ? "https://mirzailhami.github.io/outlook-signature-add-ins" : "https://localhost:3000");
              return content.toString().replace(/\${ASSET_BASE_URL}/g, assetBaseUrl);
            },
          },
          { from: ".nojekyll", to: ".nojekyll" },
          { from: "src/index.html", to: "index.html" },
          { from: "assets", to: "assets" },
          {
            from: "src/well-known",
            to: ".well-known",
            transform(content, path) {
              if (path.endsWith("microsoft-officeaddins-allowed.json")) {
                const assetBaseUrl =
                  process.env.ASSET_BASE_URL ||
                  (isProduction ? "https://mirzailhami.github.io/outlook-signature-add-ins" : "https://localhost:3000");
                const allowed = [`${assetBaseUrl}/commands.js`];
                return JSON.stringify({ allowed }, null, 2);
              }
              return content;
            },
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

    optimization: {
      splitChunks: false,
    },

    performance: {
      hints: isProduction ? "warning" : false,
      maxAssetSize: 1024 * 1024,
      maxEntrypointSize: 1024 * 1024,
    },
  };
};
