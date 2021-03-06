const path = require("path");
const TerserPlugin = require("terser-webpack-plugin");

module.exports = {
  mode: "production",
  entry: {
    index: "./src/index.js",
    "index.min": "./src/index.js"
  },
  output: {
    path: path.resolve(__dirname, "dist"),
    filename: "[name].js",
    library: "JzExcel",
    libraryExport: "default",
    globalObject: "this",
    libraryTarget: "umd"
  },
  externals: {
    // exceljs: "exceljs",
    // jszip: "jszip"
  },
  module: {
    rules: [
      {
        test: /\.js$/,
        exclude: /(node_modules|bower_components)/,
        loader: "babel-loader"
      }
    ]
  },
  resolve: {
    fallback: { stream: false }
  },
  optimization: {
    minimize: true,
    minimizer: [new TerserPlugin({ include: /\.min\.js$/ })]
  }
};
