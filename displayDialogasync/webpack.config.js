const devCerts = require("office-addin-dev-certs");
const CleanWebpackPlugin = require("clean-webpack-plugin");
const CopyWebpackPlugin = require("copy-webpack-plugin");
const CustomFunctionsMetadataPlugin = require("custom-functions-metadata-plugin");
const HtmlWebpackPlugin = require("html-webpack-plugin");
const fs = require("fs");
const webpack = require("webpack");

module.exports = async (env, options) => {
  const dev = options.mode === "development";
  const config = {
    devtool: "source-map",
    entry: {
      functions: "./src/functions/functions.ts",
      polyfill: "@babel/polyfill",
      taskpane: "./src/taskpane/taskpane.ts",
      commands: "./src/commands/commands.ts"
      // auth:"./src/auth/auth.ts",
      // AuthService: "./src/AuthService/AuthService.ts"
      // AuthSigninSerivce:"./src/AuthSigninService/AuthSigninService.ts"
    },
    resolve: {
      extensions: [".ts", ".tsx", ".html", ".js"]
    },
    module: {
      rules: [
        {
          test: /\.ts$/,
          exclude: /node_modules/,
          use: "babel-loader"
        },
        {
          test: /\.tsx?$/,
          exclude: /node_modules/,
          use: "ts-loader"
        },
        {
          test: /\.html$/,
          exclude: /node_modules/,
          use: "html-loader"
        },
        {
          test: /\.(png|jpg|jpeg|gif)$/,
          use: "file-loader"
        }
      ]
    },
    plugins: [
      new CleanWebpackPlugin({
        cleanOnceBeforeBuildPatterns: dev ? [] : ["**/*"]
      }),
      new CustomFunctionsMetadataPlugin({
        output: "functions.json",
        input: "./src/functions/functions.ts"
      }),
      new HtmlWebpackPlugin({
        filename: "functions.html",
        template: "./src/functions/functions.html",
        chunks: ["polyfill", "functions"]
      }),
      new HtmlWebpackPlugin({
        filename: "Dialog.html",
        template: "./src/Dialog/Dialog.html",
        chunks: ["polyfill", "Dialog"]
      }),
      new HtmlWebpackPlugin({
        filename: "taskpane.html",
        template: "./src/taskpane/taskpane.html",
        chunks: ["polyfill", "taskpane"]
      }),
      new CopyWebpackPlugin([
        { 
          to: "Login-16.png",          
        from: "./src/assets/Login-16.png"
      }       
      ]),
      new CopyWebpackPlugin([
        { 
          to: "Login-32.png",          
        from: "./src/assets/Login-32.png"
      }       
      ]),
      new CopyWebpackPlugin([
        { 
          to: "Login-80.png",          
        from: "./src/assets/Login-80.png"
      }       
      ]),
      new CopyWebpackPlugin([
        {
          to: "taskpane.css",
          from: "./src/taskpane/taskpane.css"
        }
      ]),
      new HtmlWebpackPlugin({
        filename: "commands.html",
        template: "./src/commands/commands.html",
        chunks: ["polyfill", "commands"]
      }),
      // new HtmlWebpackPlugin({
      //   filename: "AuthService.html",        
      //   template: "./src/AuthService/AuthService.html",
      //   chunks: ["polyfill", "AuthService"]
      // }),
      // new HtmlWebpackPlugin({
      //   filename: "AuthSigninService.html",        
      //   template: "./src/AuthSigninService/AuthSigninService.html",
      //   chunks: ["polyfill", "AuthSigninService"]
      // }),
      new HtmlWebpackPlugin({
        filename: "auth.html",        
        template: "./src/auth/auth.html",
        chunks: ["polyfill", "auth"]
      })
    ],
    devServer: {
      headers: {
        "Access-Control-Allow-Origin": "*"
      },
      https: await devCerts.getHttpsServerOptions(),
      port: 3000
    }
  };

  return config;
};
