//Dependencias
const xlsx = require( 'node-xlsx');
const fs = require('fs')
const mysql = require('mysql')
 
//Cierra dependencias
//MYSQL
const connection = mysql.createConnection({
    host: 'localhost',
    user: 'colo13',
    password: '123',
    database: 'colo13v'
})
 
const pool = mysql.createPool({
    host: 'localhost',
    user: 'colo13',
    password: '123',
    database: 'colo13v'
})
connection.connect((err) => {if(err) throw err;
console.log('Database conectada')})
//Cerrar MYSQL
//Auxiliares
auxdatos = '';
//Cierra auxiliares
 
function incertar(){
    let insertQuery = 'INSERT INTO miembros (usuario, contrase, membresia) VALUES (?, ?, ?)'
    let query = mysql.format(insertQuery,['colo13', 'picadadetomate', 'comun'])
    pool.getConnection(function(err, connection){if(err) throw err;
        connection.query(query, function(err,result){
            if(err) throw err;
            console.log(result, 'paso')
            connection.release()
        })
    })
}
function leer(){
    pool.getConnection(function(err, connection){if(err) throw err;
    connection.query('SELECT * FROM miembros', function(err,result) {if(err) throw err; console.log(result)
 
    console.log(result[0].contrase)
 
    })
    connection.release()
})
}
function actualizar(){
    pool.getConnection(function(err, connection){if(err) throw err;
        let updatequery = 'UPDATE miembros SET usuario = ? WHERE ID = ?';
        let query = mysql.format(updatequery,['kkopz', '1'])
    connection.query(query, function(err,result) {if(err) throw err; console.log(result)
 
    })
    connection.release()
})
}
function creartabla(){
    pool.getConnection(function(err, connection){if(err) throw err;
        connection.query('CREATE TABLE miembros (ID int NOT NULL AUTO_INCREMENT,usuario varchar(255) NOT NULL,contrase varchar(255),membresia varchar(255),PRIMARY KEY (ID));', function(err,result) {if(err) throw err; console.log(result)})
        connection.release()
    })
}
 
function borrardato(){
    pool.getConnection(function(err, connection){if(err) throw err;
        let removequery = 'DELETE FROM miembros WHERE ID = ?';
        let query = mysql.format(removequery,['1'])
    connection.query(query, function(err,result) {if(err) throw err; console.log(result)
 
    })
    connection.release()
})
}
 
 
function importar(ruta){
    usuariosimportar = []
    contraseñasimportar = []
    aux = xlsx.parse(ruta)
    console.log(aux[0].data[0][0],':')
for(i=1;i< aux[0].data.length;i++){
    if(aux[0].data[i][0] === undefined || aux[0].data[i][0]===[]) continue;
   console.log(aux[0].data[i][0],aux[0].data[0][2]+ ': ', aux[0].data[i][2])
   usuariosimportar.push(aux[0].data[i][0])
   contraseñasimportar.push(aux[0].data[i][2])
}
// console.log(usuariosimportar)
// console.log(contraseñasimportar)
// console.log(aux[0].data)
exportar(usuariosimportar,contraseñasimportar)
// fs.writeFileSync('./exportadoff.xlsx', datoaexp , 'binary');
}
 
function exportar(usuarios,contraseñas){
    data = []
    data.push(['    Usuarios    ','ID','    Contraseñas   '])
    data.push([]) 
    for(k=0;k<usuarios.length;k++){data.push([usuarios[k],k,contraseñas[k]])}
    console.log(data)
    const sheetOptions = {'!cols': [{wch: 30}, {wch: 6}, {wch: 30}, {wch: 7}]};
 
    datoaexp = xlsx.build([{name: 'cuentas', data: data}], {sheetOptions})
    fs.writeFileSync('./cuentas.xlsx', datoaexp , 'binary');
}
function prueba(){
    const rowAverage = [[{t: 'n', z: 10, f: '=AVERAGE(2:2)'}], [1, 2, 3]];
var buffer = xlsx.build([{name: 'Average Formula', data: rowAverage}]);
fs.writeFileSync('./prueba.xlsx', buffer , 'binary');
}
 
importar(`./importarexcel.xlsx`)
// prueba()
// creartabla()
// incertar()
// leer()
// actualizar()
// borrardato()
