const path = require('path');
const CopyWebpackPlugin = require('copy-webpack-plugin');
const webpack = require('webpack');

// Generate build timestamp
const buildTime = new Date().toISOString().replace('T', ' ').substring(0, 19) + ' UTC';

module.exports = {
  entry: {
    events: './src/events.ts',
    taskpane: './src/taskpane.ts'
  },
  output: {
    path: path.resolve(__dirname, 'dist'),
    filename: '[name].js'
  },
  resolve: {
    extensions: ['.ts', '.js']
  },
  module: {
    rules: [
      {
        test: /\.ts$/,
        use: 'ts-loader',
        exclude: /node_modules/
      }
    ]
  },
  devServer: {
    static: {
      directory: path.join(__dirname, 'dist')
    },
    compress: true,
    port: 3000,
    server: 'https',
    hot: true,
    headers: {
      "Access-Control-Allow-Origin": "*"
    }
  },
  plugins: [
    new webpack.DefinePlugin({
      '__BUILD_TIME__': JSON.stringify(buildTime)
    }),
    new CopyWebpackPlugin({
      patterns: [
        { from: 'public/taskpane.html', to: '.' },
        { from: 'public/runtime.html', to: '.' },
        { from: 'public/assets', to: 'assets' }
        // NOTE: manifest.xml is NOT copied to dist
        // It goes to Microsoft 365 Admin Center > Integrated Apps
      ]
    })
  ]
};
