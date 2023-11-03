const express = require('express');
const app = express();
const { config } = require('./config/config');
const cors = require('cors');
const steinConnection = require('./service/post.service');
const PORT = process.env.PORT || 3000;

app.use(express.json());
app.use(cors());

app.get('/api/oc', (req, res) => {
  res.status(200).send("OK, Current path: OC");
});

app.post('/api/oc', async (req, res) => {
  try {
    const body = req.body;
    const ID = await steinConnection(body, config.steinURI, 'OC');
    res.status(200).send(`${ID}`);
  } catch (error) {
    console.log(error);
    res.status(500).json({
      msg: 'Error'
    });
  }
});

app.get('/api/os', (req, res) => {
  res.status(200).send("OK, Current path: OS");
});

app.post('/api/os', async (req, res) => {
  try {
    const body = req.body;
    const ID = await steinConnection(body, config.steinURI, 'OS');
    res.status(200).send(`${ID}`);
  } catch (error) {
    console.log(error);
    res.status(500).json({
      msg: 'Error'
    });
  }
});

app.get('/api/test', (req, res) => {
  res.status(200).send("OK, Current path: OC");
});


app.post('/api/test', async (req, res) => {

  try {
    const body = req.body;
    res.status(200).send(body);
  } catch (error) {
    res.status(500).json({
      msg: 'Error'
    });
  }
})

app.listen(PORT, () => {
  console.log(`App running on ${PORT}`);
});
