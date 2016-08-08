var loadJQuery = true;
var loadReact = true;

var webpack = require("webpack");
var path = require("path");

var vendorEntry = [];
if (loadJQuery) {
    vendorEntry.push("jquery");
}
if (loadReact) {
    vendorEntry.push("react", "react-dom");
}

module.exports = {
    cache: true,
    entry: {
        main: ["./src/js/index.tsx"],
        vendor: vendorEntry
    },
    output: {
        path: path.resolve('./release'),
        publicPath: '/release/',
        filename: "[name].bundle.js",
        chunkFilename: "[id].bundle.js",
        library: "webpack",
        libraryTarget: "window"
    },

    resolve: {
        // Add '.ts' and '.tsx' as resolvable extensions.
        extensions: ["", ".webpack.js", ".web.js", ".ts", ".tsx", ".js", ".jsx"]
    },

    module: {
        loaders: [
            { test: /\.css$/, loader: "style!css-loader" },
            // All files with a '.ts' or '.tsx' extension will be handled by 'ts-loader'.
            { test: /\.tsx?$/, include: /src/, loader: "ts-loader?silent=true" },
        ]
    },

    // // When importing a module whose path matches one of the following, just
    // // assume a corresponding global variable exists and use that instead.
    // // This is important because it allows us to avoid bundling all of our
    // // dependencies, which allows browsers to cache those libraries between builds.
    //  externals: {
    //      "react": "React",
    //      "react-dom": "ReactDOM"
    //  },

    plugins: [
        loadJQuery ? new webpack.ProvidePlugin({
            // Automtically detect jQuery and $ as free var in modules
            // and inject the jquery library
            // This is required by many jquery plugins
            jQuery: "jquery",
            $: "jquery"
        }) : new webpack.ProvidePlugin({}),

        //Remove all modules in the vendor chunk from the app chunk.
        new webpack.optimize.CommonsChunkPlugin(
            /* chunkName= */"vendor",
            /* filename= */"vendor.bundle.js")
    ]
}