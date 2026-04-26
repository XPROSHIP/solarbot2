const {join} = require('path');

module.exports = {
  // Chrome'u render sunucusunun silemeyeceği proje klasörünün içine kurmaya zorlar
  cacheDirectory: join(__dirname, '.cache', 'puppeteer'),
};