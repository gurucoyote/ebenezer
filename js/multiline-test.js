var readline = require('readline');

var input = [];
var rl = readline.createInterface({
    input: process.stdin,
    output: process.stdout,
historySize : 0,
});

rl.prompt();
rl.write("This is your first line of text, please expand.");

rl.on('line', function (text) {
    input.push(text);
});
rl.on('close', function () {
    console.log(input.join('\n'));
    process.exit(0);
});
