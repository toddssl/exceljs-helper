const path = require('path');
const HtmlWebpackPlugin = require('html-webpack-plugin');
const UglifyJsPlugin = require('uglifyjs-webpack-plugin');
const package = require('./package.json');
const version = package.version;

module.exports = {
  entry: './src/index.js',
  mode: 'development',
  devServer: {
    static: {
      directory: path.join(__dirname, ''),
    },
    // hot: true,
    compress: true,
    port: 8083,
  },  
  plugins: [
    new HtmlWebpackPlugin({
      title: 'test',
      template: './index.html'
    }),
    new UglifyJsPlugin(),
  ],  
  output: {
    // filename: '[name].[contenthash].bundle.js',
    filename: `exceljs-helper-${version}.bundle.js`,
    // filename: `[name].${process.env.RELEASE_ID}.[chunkhash:8].js`,
    // filename: `[name].${process.version}.bundle.js`,
    path: path.resolve(__dirname, 'dist'),
    library: 'exceljsHelper',
  },
};