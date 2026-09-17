/* eslint-disable no-undef */

const fs = require("fs");
const os = require("os");
const path = require("path");
const CopyWebpackPlugin = require("copy-webpack-plugin");
const HtmlWebpackPlugin = require("html-webpack-plugin");

const urlDev = "https://localhost:3000/";
// Production host (Azure Static Web Apps). Set PROD_URL in the environment/CI to the
// real site, e.g. PROD_URL=https://swe-spec-formatter.azurestaticapps.net/ — it MUST
// end with a trailing slash. The production build rewrites every localhost:3000
// manifest/asset URL to this base.
const urlProd = process.env.PROD_URL || process.env.BETA_URL || "https://REPLACE-ME.azurestaticapps.net/";

// Read the dev certificates that office-addin-dev-certs already generated in
// ~/.office-addin-dev-certs, instead of calling devCerts.getHttpsServerOptions().
// That helper shells out to PowerShell to install/verify the CA, which fails under
// Windows Constrained Language Mode. Generation succeeds; we just consume the files and
// trust the CA separately with `certutil` (a native exe). Run once, if needed:
//   npm.cmd run dev-certs   ->   certutil -user -addstore Root "%USERPROFILE%\.office-addin-dev-certs\ca.crt"
function getHttpsOptions() {
  const certDir = path.join(os.homedir(), ".office-addin-dev-certs");
  return {
    ca: fs.readFileSync(path.join(certDir, "ca.crt")),
    key: fs.readFileSync(path.join(certDir, "localhost.key")),
    cert: fs.readFileSync(path.join(certDir, "localhost.crt")),
  };
}

module.exports = async (env, options) => {
  const dev = options.mode === "development";
  const config = {
    devtool: "source-map",
    entry: {
      polyfill: ["core-js/stable", "regenerator-runtime/runtime"],
      taskpane: ["./src/taskpane/taskpane.ts", "./src/taskpane/taskpane.html"],
      commands: "./src/commands/commands.ts",
    },
    output: {
      clean: true,
    },
    resolve: {
      extensions: [".ts", ".html", ".js"],
    },
    module: {
      rules: [
        {
          test: /\.ts$/,
          exclude: /node_modules/,
          use: {
            loader: "babel-loader",
            options: {
              presets: ["@babel/preset-typescript"],
            },
          },
        },
        {
          test: /\.html$/,
          exclude: /node_modules/,
          use: "html-loader",
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
      new CopyWebpackPlugin({
        patterns: [
          {
            from: "appPackage/assets/*",
            to: "assets/[name][ext][query]",
          },
          {
            from: "appPackage/manifest*.json",
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
      headers: {
        "Access-Control-Allow-Origin": "*",
      },
      server: {
        type: "https",
        options: env.WEBPACK_BUILD || options.https !== undefined ? options.https : await getHttpsOptions(),
      },
      port: process.env.npm_package_config_dev_server_port || 3000,
    },
  };

  return config;
};
