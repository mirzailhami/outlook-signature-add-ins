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

    entry: {
      commands: [
        "core-js/stable",
        "regenerator-runtime/runtime",
        path.resolve(__dirname, "src/commands/commands.js"),
        path.resolve(__dirname, "src/commands/graph.js"),
        path.resolve(__dirname, "src/commands/helpers.js"),
      ],
    },

    output: {
      path: path.resolve(__dirname, "dist"),
      filename: "commands.js",
      clean: true,
    },

    resolve: {
      extensions: [".js", ".html"],
      fallback: { https: require.resolve("https-browserify"), http: require.resolve("stream-http") },
    },

    module: {
      rules: [
        {
          test: /\.js$/,
          include: [path.resolve(__dirname, "src"), /node_modules\/(core-js|regenerator-runtime|luxon)/],
          use: {
            loader: "babel-loader",
            options: {
              presets: [
                [
                  "@babel/preset-env",
                  {
                    targets: {
                      browsers: ["last 2 versions", "not dead", "ie 11"],
                    },
                    modules: "commonjs",
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
      new Dotenv({
        path: path.resolve(__dirname, isProduction ? ".env.production" : ".env"),
        safe: true,
        allowEmptyValues: true,
        systemvars: true,
      }),
      new HtmlWebpackPlugin({
        template: "./src/taskpane/taskpane.html",
        filename: "taskpane.html",
        chunks: ["commands"],
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
        chunks: ["commands"],
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
