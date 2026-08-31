/* eslint-disable no-undef */

const fs = require("fs");
const path = require("path");
const webpack = require("webpack");
const devCerts = require("office-addin-dev-certs");
const CopyWebpackPlugin = require("copy-webpack-plugin");
const HtmlWebpackPlugin = require("html-webpack-plugin");

// The task pane bundle is public, so .env is read through an allow-list rather than wholesale.
// The connected app's consumer secret lives in the same file for reference but must never ship
// to the browser — it is the token server's to hold. Adding a key here makes it public.
const CLIENT_ENV_KEYS = ["REACT_APP_CONSUMER_KEY", "REACT_APP_REDIRECT_URI"];

// .env is gitignored, so a fresh clone has none. Missing values resolve to "" and the sign-in
// button reports the misconfiguration itself, rather than the build failing here or Salesforce
// rejecting an "undefined" client_id later.
function readClientEnv() {
  const envPath = path.resolve(__dirname, ".env");
  const values = {};

  if (fs.existsSync(envPath)) {
    for (const line of fs.readFileSync(envPath, "utf8").split(/\r?\n/)) {
      const match = /^\s*([A-Za-z_][A-Za-z0-9_]*)\s*=\s*(.*)$/.exec(line);
      if (!match) continue; // blank lines and # comments
      values[match[1]] = match[2].trim().replace(/^['"]|['"]$/g, "");
    }
  }

  const defined = {};
  for (const key of CLIENT_ENV_KEYS) {
    defined[`process.env.${key}`] = JSON.stringify(values[key] || "");
  }
  return defined;
}

async function getHttpsOptions() {
  const httpsOptions = await devCerts.getHttpsServerOptions();
  return { ca: httpsOptions.ca, key: httpsOptions.key, cert: httpsOptions.cert };
}

const urlDev = "https://localhost:3001";
const urlProd = "https://your-server-domain.com";

module.exports = async (env, options) => {
  const dev = options.mode === "development";
  const config = {
    devtool: "source-map",
    entry: {
      polyfill: ["core-js/stable", "regenerator-runtime/runtime"],
      react: ["react", "react-dom"],
      taskpane: {
        import: ["./src/index.jsx", "./src/taskpane/taskpane.html"],
        dependOn: "react",
      },
      commands: "./src/commands/commands.js",
      authcallback: "./public/auth-callback.html"
    },
    output: {
      clean: true,
    },
    resolve: {
      extensions: [".js", ".jsx", ".html"],
    },
    module: {
      rules: [
        {
          test: /\.jsx?$/,
          use: {
            loader: "babel-loader",
          },
          exclude: /node_modules/,
        },
        {
          test: /\.html$/,
          exclude: /node_modules/,
          use: "html-loader",
        },
        {
          test: /\.(png|jpg|jpeg|ttf|woff|woff2|gif|ico)$/,
          type: "asset/resource",
          generator: {
            filename: "assets/[name][ext][query]",
          },
        },
      ],
    },
    plugins: [
      new webpack.DefinePlugin(readClientEnv()),
      new HtmlWebpackPlugin({
        filename: "taskpane.html",
        template: "./src/taskpane/taskpane.html",
        chunks: ["polyfill", "taskpane", "react"],
      }),
      new HtmlWebpackPlugin({
        filename: "auth-callback.html",
        template: "./public/auth-callback.html"
      }),
      new CopyWebpackPlugin({
        patterns: [
          {
            from: "assets/*",
            to: "assets/[name][ext][query]",
          },
          {
            from: "manifest*.xml",
            to: "[name]" + "[ext]",
            transform(content) {
              if (dev) {
                return content;
              } else {
                return content.toString().replace(new RegExp(urlDev, "g"), urlProd);
              }
            },
          },
        ],
      }),
      new HtmlWebpackPlugin({
        filename: "commands.html",
        template: "./src/commands/commands.html",
        chunks: ["polyfill", "commands"],
      }),
    ],
    devServer: {
      hot: true,
      headers: {
        "Access-Control-Allow-Origin": "*",
      },
      server: {
        type: "https",
        options: env.WEBPACK_BUILD || options.https !== undefined ? options.https : await getHttpsOptions(),
      },
      port: process.env.npm_package_config_dev_server_port || 3001,
    },
  };

  return config;
};