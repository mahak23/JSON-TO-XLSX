const { generateXLSFile } = require('./app')

const data = {
    id: 1866,
    nroProforma: "0000001818",
    fechaEmision: "2020-08-29T15:11:45.150Z",
    codigoCliente: "1194",
    estado: "R",
    observacionCliente: null,
    codigoVendedor: null,
    observacionPedido: null,
    detalleProforma: [
        {
            codigoTela: "000012",
            id: 94804,
            nroProforma: "0000001818",
            partida: "AD2235",
            descripcionTela: "RIB PESCO 24/1 1X1 2 CABOS",
            codigoColor: "62O099",
            descripcionColor: "ACERO AQP RX 13 JASPEADO PESCO",
            numeroRollo: 1,
            peso: 5.5555,
            codigoBarra: "0001744717",
            codigoVenta: "0332131",
            precioVenta: 26,
            precioSistema: 26,
        },
        {
            codigoTela: "000013",
            id: 94805,
            nroProforma: "0000001818",
            partida: "AD2235",
            descripcionTela: "SIB PESCO 24/1 1X1 2 CABOS",
            codigoColor: "62O099",
            descripcionColor: "ACERO AQP RX 13 JASPEADO PESCO",
            numeroRollo: 2,
            peso: 5.555,
            codigoBarra: "0001744718",
            codigoVenta: "0332131",
            precioVenta: 26,
            precioSistema: 26,
        },
    ],
    cliente: {
        codigoCliente: "1194",
        razonSocial: "CESAR HUANCA",
        ruc: "",
        estado: "A",
    },
};


generateXLSFile(data, "./test.xlsx");
