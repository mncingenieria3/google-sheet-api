require('dotenv').config();

const config = {
  env: process.env.NODE_ENV || 'dev',
  isProd: process.env.NODE_ENV === 'production',
  port: process.env.PORT || 3000,
  steinURI: process.env.STEIN_URI,
  apiKey: process.env.API_KEY
};

module.exports = { config }
