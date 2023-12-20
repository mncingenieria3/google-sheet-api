async function testEndpoint(body) {
  const fakeId = Math.floor(Math.random() * 100);
  return fakeId;
}

module.exports = testEndpoint;
