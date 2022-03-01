//Dependencias
const xlsx = require( 'node-xlsx');
const fs = require('fs')

//Cierra dependencias


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
