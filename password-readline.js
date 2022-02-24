const l = console.log;
const stdout = process.stdout;
const stdin = process.stdin;
const readline = require("readline");
stdout.write("password:");
stdin.setRawMode(true); // needed so that each key triggers its own data event
stdin.setEncoding("utf-8"); // needed so we get data.charCodeAt
let input = "";

const pn = (data, _) => {
  const c = data;
  switch (c) {
    case "\u0004": // Ctrl-d
    case "\r":
    case "\n":
      return enter();
    case "\u0003": // Ctrl-c
      return ctrlc();
    default:
      // backspace
      if (c.charCodeAt(0) === 8) return backspace();
      else return newchar(c);
  }
};

stdin.on("data", pn);

function enter() {
  stdin.removeListener("data", pn);
  l("\nYour password is: " + input);
  stdin.setRawMode(false);
  stdin.pause();
}

function ctrlc() {
  stdin.removeListener("data", pn);
  stdin.setRawMode(false);
  stdin.pause();
}

function newchar(c) {
  input += c;
  stdout.write("*");
}

function backspace() {
  const pslen = "password:".length;
readline.moveCursor(stdin, -1, 0);
  stdout.write(" ");
  input = input.slice(0, input.length - 1);
}
