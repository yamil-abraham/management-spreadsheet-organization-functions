/*Powered by eyamilabraham*/
/*TESORERIA GH EMBRAGUES - PARTE 1*/
function onOpen() {
    var ui = SpreadsheetApp.getUi();
    ///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////MENU
    ui.createMenu('Ordenar')

        /* .addSubMenu(SpreadsheetApp.getUi().createMenu('GENERAL')
                    .addSubMenu(SpreadsheetApp.getUi().createMenu('Ordenar por Tipo')
                               .addItem('INGRESOS X VENTA', 'sortByTypeIngresos')
                               .addItem('OTROS INGRESOS', 'sortByTypeColumnaAgregada1')
                               .addItem('PASE DE EFECTIVO A PROVEEDORES', 'sortByTypeCajaProveedores')
                               .addItem('RETRIBUCIONES', 'sortByTypeRetribuciones')
                               .addItem('SUELDOS', 'sortByTypeSueldos')
                               .addItem('CARGAS SOCIALES', 'sortByTypeCargasSociales')
                               .addItem('GASTOS OPERATIVOS', 'sortByTypeGastosOperativos')
                               .addItem('HONORARIOS Y SERVICIOS CONTRATADOS', 'sortByTypeHonorariosyServiciosContratados')
                               .addItem('COMPRA DE INSUMOS', 'sortByTypeCompradeInsumos')
                               .addItem('COMPRA DE MATERIALES', 'sortByTypeCompradeMateriales')
                               .addItem('ALQUILER, IMPUESTOS Y SERVICIOS', 'sortByTypeAlquilerImpuestosyServicios')
                               .addItem('RETIROS Y ACUERDOS', 'sortByTypeRetirosyAcuerdos')
                               .addItem('DEPOSITOS EN EFECTIVO', 'sortByTypeDepositosenEfectivo')
                               .addItem('DEVOLUCIONES DE PRESTAMOS', 'sortByTypeDevolucionesdePrestamos')
                               .addItem('PRESTAMOS OTORGADOS' , 'sortByTypePrestamosOtorgados'))

                    .addSubMenu(SpreadsheetApp.getUi().createMenu('Ordenar por Fecha')
                               .addItem('INGRESOS X VENTA', 'sortByDateIngresos')
                               .addItem('OTROS INGRESOS', 'sortByDateColumnaAgregada1')
                               .addItem('PASE DE EFECTIVO A PROVEEDORES', 'sortByDateCajaProveedores')
                               .addItem('RETRIBUCIONES', 'sortByDateRetribuciones')
                               .addItem('SUELDOS', 'sortByTypeSueldos')
                               .addItem('CARGAS SOCIALES', 'sortByDateCargasSociales')
                               .addItem('GASTOS OPERATIVOS', 'sortByDateGastosOperativos')
                               .addItem('HONORARIOS Y SERVICIOS CONTRATADOS', 'sortByDateHonorariosyServiciosContratados')
                               .addItem('COMPRA DE INSUMOS', 'sortByDateCompradeInsumos')
                               .addItem('COMPRA DE MATERIALES', 'sortByDateCompradeMateriales')
                               .addItem('ALQUILER, IMPUESTOS Y SERVICIOS', 'sortByDateAlquilerImpuestosyServicios')
                               .addItem('RETIROS Y ACUERDOS', 'sortByDateRetirosyAcuerdos')
                               .addItem('DEPOSITOS EN EFECTIVO', 'sortByDateDepositosenEfectivo')
                               .addItem('DEVOLUCIONES DE PRESTAMOS', 'sortByDateDevolucionesdePrestamos')
                               .addItem('PRESTAMOS OTORGADOS' , 'sortByDatePrestamosOtorgados')))*/

        .addSubMenu(SpreadsheetApp.getUi().createMenu('PETY CHEQUES')
            .addItem('Ordenar por Fecha de Ingreso', 'sortByInDate')
            .addItem('Ordenar por Banco', 'sortByBank')
            .addItem('Ordenar por Librador', 'sortByDrawer')
            .addItem('Ordenar por Fecha de Vencimiento', 'sortByOutDate')
            .addItem('Ordenar por Fecha de Pago' , 'sortByPaymentDate')
            .addItem('Ordenar por Destinatario' , 'sortByAdresse'))
        .addSubMenu(SpreadsheetApp.getUi().createMenu('MACRO GH')
            .addSubMenu(SpreadsheetApp.getUi().createMenu('Ordenar por Tipo')
                .addItem('CREDITOS - DEPOSITOS EN EFVO', 'sortByTypeDepositosenEfvo')
                .addItem('CREDITOS - CUPONES EN TARJETA', 'sortByTypeCuponesdeTarjetas')
                .addItem('CREDITOS POR TRANSFERENCIAS', 'sortByTypeCreditosPorTransferencias')
                .addItem('DEBITOS POR CHEQUES', 'sortByTypeDebitosPorCheques')
                .addItem('TRANSFERENCIAS', 'sortByTypeTransferencias')
                .addItem('COMISIONES Y GASTOS BANCARIOS', 'sortByTypeComisionesyGastosBancarios')
                .addItem('AFIP', 'sortByTypeAFIP')
                .addItem('DEVOLUCION PRESTAMOS', 'sortByTypeDevolucionesyPrestamos')
                .addItem('OTROS DEBITOS' , 'sortByTypeOtrosDebitos'))
            .addSubMenu(SpreadsheetApp.getUi().createMenu('Ordenar por Fecha')
                .addItem('CREDITOS - DEPOSITOS EN EFVO', 'sortByDateDepositosenEfvo')
                .addItem('CREDITOS - CUPONES EN TARJETA', 'sortByDateCuponesdeTarjetas')
                .addItem('CREDITOS POR TRANSFERENCIAS', 'sortByDateCreditosPorTransferencias')
                .addItem('DEBITOS POR CHEQUES', 'sortByDateDebitosPorCheques')
                .addItem('TRANSFERENCIAS', 'sortByDateTransferencias')
                .addItem('COMISIONES Y GASTOS BANCARIOS', 'sortByDateComisionesyGastosBancarios')
                .addItem('AFIP', 'sortByDateAFIP')
                .addItem('DEVOLUCION PRESTAMOS', 'sortByDateDevolucionesyPrestamos')
                .addItem('OTROS DEBITOS' , 'sortByDateOtrosDebitos')))
        .addSubMenu(SpreadsheetApp.getUi().createMenu('MACRO CHS Y AFIP PTES')
            .addSubMenu(SpreadsheetApp.getUi().createMenu('Ordenar por Tipo')
                .addItem('CHEQUES PENDIENTES - NUEVA ERA', 'sortByTypeChequesPendientesNuevaEra')
                .addItem('CHEQUES PENDIENTES', 'sortByTypeChequesPendientes')
                .addItem('PLANES DE AFIP', 'sortByTypePlanesdeAfip')
                .addItem('AFIP - 60 CUOTAS - APORTES SEG. SOCIAL', 'sortByTypeAfip60CuotasASS')
                .addItem('AFIP - 60 CUOTAS - IVA Y CONTRA SEG SOCIAL', 'sortByTypeAfip60CuotasIyCSS')
                .addItem('GH - C. E IND MUNIC', 'sortByTypeGHCeIM')
                .addItem('MONOTRIBUTO JOSE', 'sortByTypeMonotributoJose')
                .addItem('OTROS', 'sortByTypeOtros')
                .addItem('OTROS DEBITOS' , 'sortByTypeOtrosDebitos'))
            .addSubMenu(SpreadsheetApp.getUi().createMenu('Ordenar por Tipo')
                .addItem('CHEQUES PENDIENTES - NUEVA ERA', 'sortByDateChequesPendientesNuevaEra')
                .addItem('CHEQUES PENDIENTES', 'sortByDateChequesPendientes')
                .addItem('PLANES DE AFIP', 'sortByDatePlanesdeAfip')
                .addItem('AFIP - 60 CUOTAS - APORTES SEG. SOCIAL', 'sortByDateAfip60CuotasASS')
                .addItem('AFIP - 60 CUOTAS - IVA Y CONTRA SEG SOCIAL', 'sortByDateAfip60CuotasIyCSS')
                .addItem('GH - C. E IND MUNIC', 'sortByDateGHCeIM')
                .addItem('MONOTRIBUTO JOSE', 'sortByDateMonotributoJose')
                .addItem('OTROS', 'sortByDateOtros')))

        /* .addSubMenu(SpreadsheetApp.getUi().createMenu('St RIO')
                    .addSubMenu(SpreadsheetApp.getUi().createMenu('Ordenar por Tipo')
                               .addItem('CREDITOS' , 'sortByTypeCreditosStRIO')
                               .addItem('DEBITOS GH' , 'sortByTypeDebitosGH')
                               .addItem('DEBITOS JOSE' , 'sortByTypeDebitosJose'))
                    .addSubMenu(SpreadsheetApp.getUi().createMenu('Ordenar por Fecha')
                               .addItem('CREDITOS' , 'sortByDateCreditosStRIO')
                               .addItem('DEBITOS GH' , 'sortByDateDebitosGH')
                               .addItem('DEBITOS JOSE' , 'sortByDateDebitosJose')))*/
        .addToUi()
}
//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////FUNCTIONS
/*------------------------------------------------Ingresos------------------------------------------------*/
/*function sortByTypeIngresos() {
  var  ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheets()[2];
  var range = sheet.getRange("A14:C800");
  range.sort(1);

}
function sortByDateIngresos() {
  var  ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheets()[2];
  var range = sheet.getRange("A14:C800");
  range.sort({column: 2, ascending: true});

}*/
/*------------------------------------------------Columna Agregada 1------------------------------------------------*/
/*function sortByTypeColumnaAgregada1() {
  var  ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheets()[2];
  var range = sheet.getRange("D14:F800");
  range.sort(4);

}
function sortByDateColumnaAgregada1() {
  var  ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheets()[2];
  var range = sheet.getRange("D14:F800");
  range.sort({column: 5, ascending: true});

}*/
/*------------------------------------------------Columna Agregada 2------------------------------------------------*/
/*function sortByTypeColumnaAgregada2() {
  var  ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheets()[2];
  var range = sheet.getRange("G14:I800");
  range.sort(7);

}
function sortByDateColumnaAgregada2() {
  var  ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheets()[2];
  var range = sheet.getRange("G14:I800");
  range.sort({column: 8, ascending: true});

}*/
/*------------------------------------------------Caja Proveedores------------------------------------------------*/
/*function sortByTypeCajaProveedores() {
  var  ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheets()[2];
  var range = sheet.getRange("J14:L800");
  range.sort(10);

}
function sortByDateCajaProveedores() {
  var  ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheets()[2];
  var range = sheet.getRange("J14:L800");
  range.sort({column: 11, ascending: true});

}*/
/*------------------------------------------------Retribuciones------------------------------------------------*/
/*function sortByTypeRetribuciones() {
  var  ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheets()[2];
  var range = sheet.getRange("M14:O800");
  range.sort(13);

}
function sortByDateRetribuciones() {
  var  ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheets()[2];
  var range = sheet.getRange("M14:O800");
  range.sort({column: 14, ascending: true});

}*/
/*------------------------------------------------Sueldos------------------------------------------------*/
/*function sortByTypeSueldos() {
  var  ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheets()[2];
  var range = sheet.getRange("P14:R800");
  range.sort(16);

}
function sortByDateSueldos() {
  var  ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheets()[2];
  var range = sheet.getRange("P14:R800");
  range.sort({column: 17, ascending: true});

}*/
/*------------------------------------------------Cargas Sociales------------------------------------------------*/
/*function sortByTypeCargasSociales() {
  var  ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheets()[2];
  var range = sheet.getRange("S14:U800");
  range.sort(19);

}
function sortByDateCargasSociales() {
  var  ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheets()[2];
  var range = sheet.getRange("S14:U800");
  range.sort({column: 20, ascending: true});

}*/
/*------------------------------------------------Gastos Operativos------------------------------------------------*/
/*function sortByTypeGastosOperativos() {
  var  ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheets()[2];
  var range = sheet.getRange("V14:X800");
  range.sort(22);

}
function sortByDateGastosOperativos() {
  var  ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheets()[2];
  var range = sheet.getRange("V14:X800");
  range.sort({column: 23, ascending: true});

}*/
/*------------------------------------------------Honorarios y Servicios Contratados------------------------------------------------*/
/*function sortByTypeHonorariosyServiciosContratados() {
  var  ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheets()[2];
  var range = sheet.getRange("Y14:AA800");
  range.sort(25);

}
function sortByDateHonorariosyServiciosContratados() {
  var  ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheets()[2];
  var range = sheet.getRange("Y14:AA800");
  range.sort({column: 26, ascending: true});

}*/
/*------------------------------------------------Compra de Insumos------------------------------------------------*/
/*function sortByTypeCompradeInsumos() {
  var  ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheets()[2];
  var range = sheet.getRange("AB14:AD800");
  range.sort(28);

}
function sortByDateCompradeInsumos() {
  var  ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheets()[2];
  var range = sheet.getRange("AB14:AD800");
  range.sort({column: 29, ascending: true});

}*/
/*------------------------------------------------Compra de Materiales------------------------------------------------*/
/*function sortByTypeCompradeMateriales() {
  var  ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheets()[2];
  var range = sheet.getRange("AE14:AG800");
  range.sort(31);

}
function sortByDateCompradeMateriales() {
  var  ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheets()[2];
  var range = sheet.getRange("AE14:AG800");
  range.sort({column: 32, ascending: true});

}*/
/*------------------------------------------------Alquiler, Impuestos y Servicios------------------------------------------------*/
/*function sortByTypeAlquilerImpuestosyServicios() {
  var  ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheets()[2];
  var range = sheet.getRange("AH14:AJ800");
  range.sort(34);

}
function sortByDateAlquilerImpuestosyServicios() {
  var  ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheets()[2];
  var range = sheet.getRange("AH14:AJ800");
  range.sort({column: 35, ascending: true});

}*/
/*------------------------------------------------Retiros y Acuerdos------------------------------------------------*/
/*function sortByTypeRetirosyAcuerdos() {
  var  ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheets()[2];
  var range = sheet.getRange("AK14:AM800");
  range.sort(37);

}
function sortByDateRetirosyAcuerdos() {
  var  ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheets()[2];
  var range = sheet.getRange("AK14:AM800");
  range.sort({column: 38, ascending: true});

}*/
/*------------------------------------------------Depositos en Efectivo------------------------------------------------*/
/*function sortByTypeDepositosenEfectivo() {
  var  ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheets()[2];
  var range = sheet.getRange("AN14:AP800");
  range.sort(40);

}
function sortByDateDepositosenEfectivo() {
  var  ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheets()[2];
  var range = sheet.getRange("AN14:AP800");
  range.sort({column: 41, ascending: true});

}*/
/*------------------------------------------------Devoluciones de Prestamos------------------------------------------------*/
/*function sortByTypeDevolucionesdePrestamos() {
  var  ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheets()[2];
  var range = sheet.getRange("AQ14:AS800");
  range.sort(43);

}
function sortByDateDevolucionesdePrestamos() {
  var  ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheets()[2];
  var range = sheet.getRange("AQ14:AS800");
  range.sort({column: 44, ascending: true});

}*/
/*------------------------------------------------Prestamos Otorgados------------------------------------------------*/
/*function sortByTypePrestamosOtorgados() {
  var  ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheets()[2];
  var range = sheet.getRange("AT14:AV800");
  range.sort(46);

}
function sortByDatePrestamosOtorgados() {
  var  ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheets()[2];
  var range = sheet.getRange("AT14:AV800");
  range.sort({column: 47, ascending: true});

}*/
/*------------------------------------------------Columna Agregada 4------------------------------------------------*/
/*function sortByTypeColumnaAgregada4() {
  var  ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheets()[2];
  var range = sheet.getRange("AW14:AY800");
  range.sort(49);

}
function sortByDateColumnaAgregada4() {
  var  ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheets()[2];
  var range = sheet.getRange("AW14:AY800");
  range.sort({column: 50, ascending: true});

}*/
/*------------------------------------------------Columna Agregada 5------------------------------------------------*/
/*function sortByTypeColumnaAgregada5() {
  var  ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheets()[2];
  var range = sheet.getRange("AZ14:BB800");
  range.sort(52);

}
function sortByDateColumnaAgregada5() {
  var  ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheets()[2];
  var range = sheet.getRange("AZ14:BB800");
  range.sort({column: 53, ascending: true});

}*/
/*------------------------------------------------Columna Agregada 6------------------------------------------------*/
/*function sortByTypeColumnaAgregada6() {
  var  ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheets()[2];
  var range = sheet.getRange("BC14:BE800");
  range.sort(55);

}
function sortByDateColumnaAgregada6() {
  var  ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheets()[2];
  var range = sheet.getRange("BC14:BE800");
  range.sort({column: 56, ascending: true});

}*/
/*------------------------------------------------CHEQUES------------------------------------------------*/
function sortByInDate() {
    var  ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheets()[21];
    var range = sheet.getRange("A8:J2700");
    range.sort(1);

}
function sortByBank() {
    var  ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheets()[21];
    var range = sheet.getRange("A8:J2700");
    range.sort({column: 2, ascending: true});

}
function sortByDrawer() {
    var  ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheets()[21];
    var range = sheet.getRange("A8:J2700");
    range.sort(4);

}
function sortByOutDate() {
    var  ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheets()[21];
    var range = sheet.getRange("A8:J2700");
    range.sort({column: 6, ascending: true});

}
function sortByPaymentDate() {
    var  ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheets()[21];
    var range = sheet.getRange("A8:J2700");
    range.sort({column: 8, ascending: true});

}

function sortByAdresse() {
    var  ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheets()[21];
    var range = sheet.getRange("A8:J2700");
    range.sort({column: 5, ascending: true});

}

/*------------------------------------------------MACRO GH------------------------------------------------*/

/*------------------------------------------------Creditos - Depositos en Efvo------------------------------------------------*/
function sortByTypeDepositosenEfvo() {
    var  ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheets()[22];
    var range = sheet.getRange("A14:C800");
    range.sort(1);

}
function sortByDateDepositosenEfvo() {
    var  ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheets()[22];
    var range = sheet.getRange("A14:C800");
    range.sort({column: 2, ascending: true});

}
/*------------------------------------------------Depositos - Cupones de Tarjeta------------------------------------------------*/
function sortByTypeCuponesdeTarjetas() {
    var  ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheets()[22];
    var range = sheet.getRange("D14:F800");
    range.sort(4);

}
function sortByDateCuponesdeTarjetas() {
    var  ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheets()[22];
    var range = sheet.getRange("D14:F800");
    range.sort({column: 5, ascending: true});

}
/*------------------------------------------------Creditos Por Transferencias------------------------------------------------*/
function sortByTypeCreditosPorTransferencias() {
    var  ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheets()[22];
    var range = sheet.getRange("G14:I800");
    range.sort(7);

}
function sortByDateCreditosPorTransferencias() {
    var  ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheets()[22];
    var range = sheet.getRange("G14:I800");
    range.sort({column: 8, ascending: true});

}
/*------------------------------------------------Debitos Por Cheques------------------------------------------------*/
function sortByTypeDebitosPorCheques() {
    var  ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheets()[22];
    var range = sheet.getRange("J14:L800");
    range.sort(10);

}
function sortByDateDebitosPorCheques() {
    var  ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheets()[22];
    var range = sheet.getRange("J14:L800");
    range.sort({column: 11, ascending: true});

}
/*------------------------------------------------Transferencias------------------------------------------------*/
function sortByTypeTransferencias() {
    var  ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheets()[22];
    var range = sheet.getRange("M14:O800");
    range.sort(13);

}
function sortByDateTransferencias() {
    var  ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheets()[22];
    var range = sheet.getRange("M14:O800");
    range.sort({column: 14, ascending: true});

}
/*------------------------------------------------Comisiones y Gastos Bancarios------------------------------------------------*/
function sortByTypeComisionesyGastosBancarios() {
    var  ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheets()[22];
    var range = sheet.getRange("P14:R800");
    range.sort(16);

}
function sortByDateComisionesyGastosBancarios() {
    var  ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheets()[22];
    var range = sheet.getRange("P14:R800");
    range.sort({column: 17, ascending: true});

}
/*------------------------------------------------AFIP------------------------------------------------*/
function sortByTypeAFIP() {
    var  ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheets()[22];
    var range = sheet.getRange("S14:U800");
    range.sort(19);

}
function sortByDateAFIP() {
    var  ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheets()[22];
    var range = sheet.getRange("S14:U800");
    range.sort({column: 20, ascending: true});

}
/*------------------------------------------------Devoluciones y Prestamos------------------------------------------------*/
function sortByTypeDevolucionesyPrestamos() {
    var  ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheets()[22];
    var range = sheet.getRange("V14:X800");
    range.sort(22);

}
function sortByDateDevolucionesyPrestamos() {
    var  ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheets()[22];
    var range = sheet.getRange("V14:X800");
    range.sort({column: 23, ascending: true});

}
/*------------------------------------------------OTROS DEBITOS------------------------------------------------*/
function sortByTypeOtrosDebitos() {
    var  ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheets()[22];
    var range = sheet.getRange("Y14:AA800");
    range.sort(25);

}
function sortByDateOtrosDebitos() {
    var  ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheets()[22];
    var range = sheet.getRange("Y14:AA800");
    range.sort({column: 26, ascending: true});

}
/*------------------------------------------------C2------------------------------------------------*/
function sortByTypeC2() {
    var  ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheets()[22];
    var range = sheet.getRange("AB14:AD800");
    range.sort(28);

}
function sortByDateC2() {
    var  ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheets()[22];
    var range = sheet.getRange("AB14:AD800");
    range.sort({column: 29, ascending: true});

}
/*------------------------------------------------C3------------------------------------------------*/
function sortByTypeC3() {
    var  ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheets()[22];
    var range = sheet.getRange("AE14:AG800");
    range.sort(31);

}
function sortByDateC3() {
    var  ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheets()[22];
    var range = sheet.getRange("AE14:AG800");
    range.sort({column: 32, ascending: true});

}
/*------------------------------------------------C4------------------------------------------------*/
function sortByTypeC4() {
    var  ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheets()[22];
    var range = sheet.getRange("AH14:AJ800");
    range.sort(34);

}
function sortByDateC4() {
    var  ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheets()[22];
    var range = sheet.getRange("AH14:AJ800");
    range.sort({column: 35, ascending: true});

}
/*------------------------------------------------C5------------------------------------------------*/
function sortByTypeC5() {
    var  ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheets()[22];
    var range = sheet.getRange("AK14:AM800");
    range.sort(37);

}
function sortByDateC5() {
    var  ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheets()[22];
    var range = sheet.getRange("AK14:AM800");
    range.sort({column: 38, ascending: true});

}
/*------------------------------------------------C6------------------------------------------------*/
function sortByTypeC6() {
    var  ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheets()[22];
    var range = sheet.getRange("AN14:AP800");
    range.sort(40);

}
function sortByDateC6() {
    var  ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheets()[22];
    var range = sheet.getRange("AN14:AP800");
    range.sort({column: 41, ascending: true});

}
/*------------------------------------------------C7------------------------------------------------*/
function sortByTypeC7() {
    var  ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheets()[22];
    var range = sheet.getRange("AQ14:AS800");
    range.sort(43);

}
function sortByDateC7() {
    var  ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheets()[22];
    var range = sheet.getRange("AQ14:AS800");
    range.sort({column: 44, ascending: true});

}
/*------------------------------------------------C8------------------------------------------------*/
function sortByTypeC8() {
    var  ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheets()[22];
    var range = sheet.getRange("AT14:AV800");
    range.sort(46);

}
function sortByDateC8() {
    var  ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheets()[22];
    var range = sheet.getRange("AT14:AV800");
    range.sort({column: 47, ascending: true});

}
/*------------------------------------------------C9------------------------------------------------*/
function sortByTypeC9() {
    var  ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheets()[22];
    var range = sheet.getRange("AW14:AY800");
    range.sort(49);

}
function sortByDateC9() {
    var  ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheets()[22];
    var range = sheet.getRange("AW14:AY800");
    range.sort({column: 50, ascending: true});

}
/*------------------------------------------------C10------------------------------------------------*/
function sortByTypeC10() {
    var  ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheets()[22];
    var range = sheet.getRange("AZ14:BB800");
    range.sort(52);

}
function sortByDateC10() {
    var  ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheets()[22];
    var range = sheet.getRange("AZ14:BB800");
    range.sort({column: 53, ascending: true});

}
/*------------------------------------------------C11------------------------------------------------*/
function sortByTypeC11() {
    var  ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheets()[22];
    var range = sheet.getRange("BC14:BE800");
    range.sort(55);

}
function sortByDateC11() {
    var  ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheets()[22];
    var range = sheet.getRange("BC14:BE800");
    range.sort({column: 56, ascending: true});

}
/*------------------------------------------------MACRO CHS Y AFIP PTES------------------------------------------------*/

/*------------------------------------------------Cheques Pendientes - Nueva Era------------------------------------------------*/
function sortByTypeChequesPendientesNuevaEra() {
    var  ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheets()[23];
    var range = sheet.getRange("A14:C801");
    range.sort(1);

}
function sortByDateChequesPendientesNuevaEra() {
    var  ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheets()[23];
    var range = sheet.getRange("A14:C801");
    range.sort({column: 2, ascending: true});

}
/*------------------------------------------------Cheques Pendientes------------------------------------------------*/
function sortByTypeChequesPendientes() {
    var  ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheets()[23];
    var range = sheet.getRange("D14:F801");
    range.sort(4);

}
function sortByDateChequesPendientes() {
    var  ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheets()[23];
    var range = sheet.getRange("D14:F801");
    range.sort({column: 5, ascending: true});

}
/*------------------------------------------------Planes de AFIP------------------------------------------------*/
function sortByTypePlanesdeAfip() {
    var  ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheets()[23];
    var range = sheet.getRange("G14:I801");
    range.sort(7);

}
function sortByDatePlanesdeAfip() {
    var  ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheets()[23];
    var range = sheet.getRange("G14:I801");
    range.sort({column: 8, ascending: true});

}
/*------------------------------------------------Afip - plan 60 cuotas - Aportes Seg Social ------------------------------------------------*/
function sortByTypeAfip60CuotasASS() {
    var  ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheets()[23];
    var range = sheet.getRange("J14:L801");
    range.sort(10);

}
function sortByDateAfip60CuotasASS() {
    var  ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheets()[23];
    var range = sheet.getRange("J14:L801");
    range.sort({column: 11, ascending: true});

}
/*------------------------------------------------Afip - plan 60 cuotas - IVA Y Contr Seg Social ------------------------------------------------*/
function sortByTypeAfip60CuotasIyCSS() {
    var  ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheets()[23];
    var range = sheet.getRange("M14:O801");
    range.sort(13);

}
function sortByDateAfip60CuotasIyCSS() {
    var  ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheets()[23];
    var range = sheet.getRange("M14:O801");
    range.sort({column: 14, ascending: true});

}
/*------------------------------------------------GH - C.e Ind Munic ------------------------------------------------*/
function sortByTypeGHCeIM() {
    var  ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheets()[23];
    var range = sheet.getRange("P14:R801");
    range.sort(16);

}
function sortByDateGHCeIM() {
    var  ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheets()[23];
    var range = sheet.getRange("P14:R801");
    range.sort({column: 17, ascending: true});

}
/*------------------------------------------------Monotributo Jose------------------------------------------------*/
function sortByTypeMonotributoJose() {
    var  ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheets()[23];
    var range = sheet.getRange("S14:U801");
    range.sort(19);

}
function sortByDateMonotributoJose() {
    var  ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheets()[23];
    var range = sheet.getRange("S14:U801");
    range.sort({column: 20, ascending: true});

}
/*------------------------------------------------OTROS------------------------------------------------*/
function sortByTypeOtros() {
    var  ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheets()[23];
    var range = sheet.getRange("V14:X801");
    range.sort(22);

}
function sortByDateOtros() {
    var  ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheets()[23];
    var range = sheet.getRange("V14:X801");
    range.sort({column: 23, ascending: true});

}
/*------------------------------------------------CREDITOS St RIO------------------------------------------------*/
function sortByTypeCreditosStRIO() {
    var  ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheets()[26];
    var range = sheet.getRange("A14:C800");
    range.sort(1);

}
function sortByDateCreditosStRIO() {
    var  ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheets()[26];
    var range = sheet.getRange("A14:C800");
    range.sort({column: 2, ascending: true});

}
/*------------------------------------------------DEBITOS GH------------------------------------------------*/
function sortByTypeDebitosGH() {
    var  ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheets()[26];
    var range = sheet.getRange("D14:F800");
    range.sort(4);

}
function sortByDateDebitosGH() {
    var  ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheets()[26];
    var range = sheet.getRange("D14:F800");
    range.sort({column: 5, ascending: true});

}
/*------------------------------------------------DEBITOS JOSE------------------------------------------------*/
function sortByTypeDebitosJose() {
    var  ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheets()[26];
    var range = sheet.getRange("G14:I800");
    range.sort(7);

}
function sortByDateDebitosJose() {
    var  ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheets()[26];
    var range = sheet.getRange("G14:I800");
    range.sort({column: 8, ascending: true});

}
/*------------------------------------------------C12------------------------------------------------*/
function sortByTypeC12() {
    var  ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheets()[26];
    var range = sheet.getRange("J14:L800");
    range.sort(10);

}
function sortByDateC12() {
    var  ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheets()[26];
    var range = sheet.getRange("J14:L800");
    range.sort({column: 11, ascending: true});

}
/*------------------------------------------------C13------------------------------------------------*/
function sortByTypeC13() {
    var  ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheets()[26];
    var range = sheet.getRange("M14:O800");
    range.sort(13);

}
function sortByDateC13() {
    var  ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheets()[26];
    var range = sheet.getRange("M14:O800");
    range.sort({column: 14, ascending: true});

}
/*------------------------------------------------C14------------------------------------------------*/
function sortByTypeC14() {
    var  ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheets()[26];
    var range = sheet.getRange("P14:R800");
    range.sort(16);

}
function sortByDateC14() {
    var  ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheets()[26];
    var range = sheet.getRange("P14:R800");
    range.sort({column: 17, ascending: true});

}
/*------------------------------------------------C15------------------------------------------------*/
function sortByTypeC15() {
    var  ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheets()[26];
    var range = sheet.getRange("S14:U800");
    range.sort(19);

}
function sortByDateC15() {
    var  ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheets()[26];
    var range = sheet.getRange("S14:U800");
    range.sort({column: 20, ascending: true});


}
/*------------------------------------------------C16------------------------------------------------*/
function sortByTypeC16() {
    var  ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheets()[26];
    var range = sheet.getRange("V14:X800");
    range.sort(22);

}
function sortByDateC16() {
    var  ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheets()[26];
    var range = sheet.getRange("V14:X800");
    range.sort({column: 23, ascending: true});

}