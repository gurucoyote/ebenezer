const strkey = require("stringify-key");
const readline = require("readline");

const rl = readline.createInterface({
  input: process.stdin,
  output: process.stdout,
  historySize: 0,
});
rl.input.on("keypress", (_, key) => {
  console.log(key);
  console.log(strkey(key));
});
