const path = require('path');
const HtmlWebpackPlugin = require('html-webpack-plugin');
const CopyWebpackPlugin = require('copy-webpack-plugin');
const MiniCssExtractPlugin = require('mini-css-extract-plugin');

module.exports = {
    mode: 'production',
    context: path.resolve(__dirname),
    entry: './src/taskpane.js',
    output: {
        path: path.resolve(__dirname, 'dist'),
        filename: 'taskpane.js',
        clean: true,
        library: { type: 'umd' },
    },
    module: {
        rules: [
            {
                test: /\.css$/i,
                use: [
                    MiniCssExtractPlugin.loader,
                    { loader: 'css-loader', options: { sourceMap: true } },
                ],
            },
            {
                test: /\.html$/i,
                loader: 'html-loader',
                options: { minimize: false, sources: false },
            },
            {
                test: /\.js$/i,
                exclude: /node_modules/,
                use: {
                    loader: 'babel-loader',
                    options: { presets: ['@babel/preset-env'] },
                },
            },
        ],
    },
    plugins: [
        new HtmlWebpackPlugin({
            template: './src/taskpane.html',
            filename: 'taskpane.html',
        }),
        new CopyWebpackPlugin({
            patterns: [
                { from: 'assets/favicon.ico', to: 'assets/favicon.ico' },
                { from: 'assets/icon-16.png', to: 'assets/icon-16.png' },
                { from: 'assets/icon-32.png', to: 'assets/icon-32.png' },
                { from: 'assets/icon-80.png', to: 'assets/icon-80.png' },
                { from: 'manifest.xml', to: 'manifest.xml' },
                { from: 'config/xai_config.json', to: 'xai_config.json' },
            ],
        }),
        new MiniCssExtractPlugin({
            filename: 'taskpane.css',
        }),
    ],
    devServer: {
        static: path.resolve(__dirname, 'dist'),
        port: 3001,
        headers: {
            'Access-Control-Allow-Origin': '*',
        },
    },
    devtool: 'source-map',
};