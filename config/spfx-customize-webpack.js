module.exports = function(webpackConfig, taskSession, heftConfiguration, webpack) {
  webpackConfig.resolve = webpackConfig.resolve || {};
  webpackConfig.resolve.fallback = {
    ...webpackConfig.resolve.fallback,
    fs: false,
    https: false,
    http: false,
    url: false,
    stream: false,
    buffer: false,
    os: false,
    path: false
  };

  if (webpack && webpack.NormalModuleReplacementPlugin) {
    webpackConfig.plugins = webpackConfig.plugins || [];
    webpackConfig.plugins.push(
      new webpack.NormalModuleReplacementPlugin(/^node:(.*)/, (resource) => {
        resource.request = resource.request.replace(/^node:/, '');
      })
    );
  }

  return webpackConfig;
};
