const express = require('express');
const app = express();
const { config } = require('./config/config');
const cors = require('cors');
const steinConnection = require('./service/post.service');
const PORT = process.env.PORT || 3000;

app.use(express.json());
app.use(cors());

app.get('/api', (req, res) => {
  res.status(200).send("OK");
});

app.post('/api', async (req, res) => {
  try {
    const body = req.body;
    console.log(config.steinURI)
    const ID = await steinConnection(body, config.steinURI);
    res.status(200).send(`${ID}`);
  } catch (error) {
    console.log(error);
    res.status(500).json({
      msg: 'Error'
    });
  }
});

app.listen(PORT, () => {
  console.log(`App running on ${PORT}`);
});
