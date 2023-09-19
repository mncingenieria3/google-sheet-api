const SteinStore = require('stein-js-client');
const config = require('../config/config');

async function steinConnection(reqBody, URI, docType) {
  const body = reqBody[0];
  const store = new SteinStore(URI);
  let lastID = 0;
  let lastOC = 0;

  const data = await store.read(docType);
  lastID = data.length + 1;
  lastOC = lastID;
  if (docType === 'OC') {
    await store.append(docType, [
      {
        ID: lastID,
        FECHA: body.FECHA,
        'NUMERO OC': lastOC,
        PROVEEDOR: body.PROVEEDOR,
        'VALOR BRUTO': body.VALOR_BRUTO,
        'N° FRA': body.N_FRA,
        VALOR: body.VALOR,
        SOLICITANTE: body.SOLICITANTE,
        'CENTRO COSTO': body.CENTRO_COSTO,
        CIUDAD: body.CIUDAD,
        OBSERVACIONES: body.OBSERVACIONES,
      },
    ]);
  } else {
    await store.append(docType, [
      {
        ID: lastID,
        FECHA: body.FECHA,
        'NUMERO OS': lastOC,
        PROVEEDOR: body.PROVEEDOR,
        'VALOR BRUTO': body.VALOR_BRUTO,
        'N° FRA': body.N_FRA,
        VALOR: body.VALOR,
        SOLICITANTE: body.SOLICITANTE,
        'CENTRO COSTO': body.CENTRO_COSTO,
        CIUDAD: body.CIUDAD,
        OBSERVACIONES: body.OBSERVACIONES,
      },
    ]);
  }
  return lastID;
}

module.exports = steinConnection;
