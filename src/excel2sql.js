const XLSX = require('xlsx');
const fs = require('fs');
var colors = require('colors');

class excel2sql {

  constructor(parameters) {
    this.header = '';
    this.sql = '';
    this.data = [];
    this.filePath = parameters.input;
    this.output = '';
    this.tableName = '';
    this.counter = 0;
    this.headerStatement = null;

    let fileName = parameters.input.split('.');
    if (parameters.output === undefined) {
      this.output = fileName[0] + '.sql';
    } else {
      this.output = parameters.output;
    }

    if (parameters.name === undefined) {
      this.tableName = fileName[0];
    } else {
      this.tableName = parameters.name
    }
  }

  read() {
      const workbook = XLSX.readFile(this.filePath, { raw: true});
      const sheetName = workbook.SheetNames[0];
      const worksheet = workbook.Sheets[sheetName];
      const range = XLSX.utils.decode_range(worksheet['!ref']);
      const totalRows = range.e.r + 1; // Número total de filas
      const totalColumns = range.e.c + 1; // Número total de columnas
  
      // Leer los datos por fila y columna
      for (let row = 0; row < totalRows; row++) {
        let tuple = [];
        for (let col = 0; col < totalColumns; col++) {
          const cellAddress = XLSX.utils.encode_cell({ r: row, c: col });
          const cell = worksheet[cellAddress];
          tuple.push(cell ? cell.w : undefined);
        }
        this.data.push(tuple);
      }
  }

  sqlHeader(array) {
    
      let insert = 'INSERT INTO ' + this.tableName + ' (';
      insert += array.join(', ');
      insert += ') VALUES';
      return insert;
    

    this.counter--;
  }

  sqlFields(array) {
    console.log(`entro sqlFiueld`)
    array = array.map(field => {
      return field === undefined
        ? 'null'
        : '\'' + field + '\'';
    });

    let insert = '(';
    insert += array.join(', ');
    insert += ')';
    return insert;
  }

  makeSentences() {
    let preSql = [];
      while (this.data[this.counter] !== undefined) {
        console.log(`this.counter ${this.counter}`);
        let residue = this.counter % 1000;
        let set = Math.trunc(this.counter /1000);
        if (residue === 0 || this.counter === 0) {
          this.header = this.sqlHeader(this.data[this.counter]);
          preSql[set] = [];
        } else {
          preSql[set].push(this.sqlFields(this.data[this.counter]));
        }
        this.counter++;
    }

    this.sql = 'BEGIN TRANSACTION;\n';
    preSql.forEach(set => {
      this.sql += this.header + '\n';
      this.sql += set.join(',\n') + ';\n';
    });
    this.sql += 'ROLLBACK;'
  }

  write() {
    fs.writeFile(this.output, this.sql, (err) => {
      if (err)
        console.log(colors.red('Error al generar el archivo ' + err));
      else {
        console.log('Proceso terminado'.green);
      }
    });
  }


  make() {
    this.read();
    this.makeSentences();
    this.write();
  }
}

module.exports = excel2sql;