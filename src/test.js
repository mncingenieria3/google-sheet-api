const steinConnection = require("./service/post.service");
const { config } = require('./config/config');

const body = {
  PROVEEDOR: '3S SERVICIOS INTEGRALES SAS',
  FECHA: '6/09/2023',
  VALOR_BRUTO: '0',
  VALOR: '0',
  N_FRA: '001'
}
console.log(config.steinURI)

steinConnection(body, config.steinURI).then((data) => console.log(data));
