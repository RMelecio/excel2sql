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

    console.log('Valor del primer parámetro:', parameters.input);
    console.log('Valor del segundo parámetro:', parameters.output);
    console.log('Valor del tercer parámetro:', parameters.name);
    
    return false;
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
    let insert = 'INSERT INTO payment_ways (';
    insert += array.join(', ');
    insert += ') VALUES';
    return insert;
  }

  sqlFields(array) {
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
    this.data.forEach((line, index) => {
      let residue = index % 1000;
      let set = Math.trunc(index /1000);

      if (residue === 0 || index === 0) {
        this.header = this.sqlHeader(line);
        preSql[set] = [];
      } else {
        preSql[set].push(this.sqlFields(line));
      }
    });

    this.sql = 'BEGIN TRANSACTION;\n';
    preSql.forEach(set => {
      this.sql += this.header + '\n';
      this.sql += set.join(',\n') + ';\n';
    });
    this.sql += 'ROLLBACK;'
  }

  write() {
    fs.writeFile("insert.sql", this.sql, (err) => {
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