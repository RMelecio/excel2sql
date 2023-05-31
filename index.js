const { hideBin } = require('yargs/helpers');
const yargs = require('yargs/yargs')(hideBin(process.argv));

const excel2sql = require('./src/excel2sql');



const argv = yargs
  .option('input', {
      alias: 'i',
      describe: 'input file',
      demandOption: true, // Indica que el parámetro es obligatorio
      type: 'string' // Tipo de dato del parámetro (en este caso, string)
  })
  .option('output', {
      alias: 'o',
      describe: 'output file',
      demandOption: false,
      type: 'string' // Tipo de dato del parámetro (en este caso, número)
  })
  .option('name', {
      alias: 'n',
      describe: 'table name to insert',
      demandOption: false,
      type: 'string' // Tipo de dato del parámetro (en este caso, booleano)
  })
  .argv;

const obj = new excel2sql(argv);
obj.make();



