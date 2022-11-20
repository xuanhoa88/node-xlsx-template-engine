const path = require("path");

module.exports = {
  mode: "production",

  target: "node",

  entry: "./src/main.js",

  output: {
    path: path.join(__dirname, "dist"),
    filename: "[name].js",
    // https://github.com/webpack/webpack/issues/1114
    library: {
      type: "commonjs2",
    },
  },

  module: {
    rules: [
      {
        test: /\.m?js$/,
        exclude: /node_modules/,
        use: {
          loader: "babel-loader",
          options: {
            cacheDirectory: true,
          },
        },
      },
    ],
  },

  /**
   * Determine the array of extensions that should be used to resolve modules.
   */
  resolve: {
    extensions: [".js"],
    modules: ["node_modules"],
  },

  /**
   * Disables webpack processing of __dirname and __filename.
   * If you run the bundle in node.js it falls back to these values of node.js.
   * https://github.com/webpack/webpack/issues/2010
   */
  node: {
    __dirname: false,
    __filename: false,
  },
};
