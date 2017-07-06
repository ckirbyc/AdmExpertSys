$(document).ready(function () {

    $.each($('.porcentajedosDecimales'), function(i, control) {
        $(control).attr("data-a-sep", ".");
        $(control).attr("data-a-dec", ",");
        $(control).attr("data-v-max", "10000000");
        $(control).attr("data-v-min", "-10000000");
        $(control).attr("data-m-dec", "2");
        $(control).attr("data-a-sign", "%");
        $(control).attr("data-p-sign", "s");
    });

    $.each($('.porcentajeUnDecimal'), function (i, control) {
        $(control).attr("data-a-sep", ".");
        $(control).attr("data-a-dec", ",");
        $(control).attr("data-v-max", "1000");
        $(control).attr("data-v-min", "-1000");
        $(control).attr("data-m-dec", "1");
        $(control).attr("data-a-sign", "%");
        $(control).attr("data-p-sign", "s");
    });

    $.each($('.dosDecimales'), function (i, control) {
        $(control).attr("data-a-sep", ".");
        $(control).attr("data-a-dec", ",");
        $(control).attr("data-v-max", "999999999999999.99");
        $(control).attr("data-m-dec", "2");
        $(control).attr("data-v-min", "-999999999999999.99");

    });

    $.each($('.cuatroDecimales'), function (i, control) {
        $(control).attr("data-a-sep", ".");
        $(control).attr("data-a-dec", ",");
        $(control).attr("data-v-max", "999999999999999.9999");
        $(control).attr("data-m-dec", "4");
        $(control).attr("data-v-min", "-999999999999999.9999");

    });

    $.each($('.dosDecimalesTotal'), function (i, control) {
        $(control).attr("data-a-sep", ".");
        $(control).attr("data-a-dec", ",");
        $(control).attr("data-v-max", "9999999999999999.99");
        $(control).attr("data-m-dec", "2");
        $(control).attr("data-v-min", "-9999999999999999.99");

    });
    $.each($('.porcentaje'), function (i, control) {
        $(control).attr("data-v-max", "1000");
        $(control).attr("data-a-sign", "%");
        $(control).attr("data-p-sign", "s");
        $(control).attr("data-a-sep", ".");
        $(control).attr("data-a-dec", ",");
    });

    $.each($('.meses'), function (i, control) {
        $(control).attr("data-v-max", "9999");
        $(control).attr("data-a-sign", " meses");
        $(control).attr("data-a-sep", ".");
        $(control).attr("data-a-dec", ",");
        $(control).attr("data-p-sign", "s");
        $(control).attr("data-m-dec", "0");
        $(control).attr("data-v-min", "0");
    });

    $.each($('.mesesDuration'), function (i, control) {
        $(control).attr("data-v-max", "9999");
        $(control).attr("data-a-sign", " a\u00F1os");
        $(control).attr("data-a-sep", ".");
        $(control).attr("data-a-dec", ",");
        $(control).attr("data-p-sign", "s");
        $(control).attr("data-m-dec", "2");
        $(control).attr("data-v-min", "0");
    });

    $.each($('.dias'), function (i, control) {
        $(control).attr("data-v-max", "9999");
        $(control).attr("data-a-sign", " dias");
        $(control).attr("data-a-sep", ".");
        $(control).attr("data-a-dec", ",");
        $(control).attr("data-p-sign", "s");
        $(control).attr("data-m-dec", "0");
        $(control).attr("data-v-min", "0");
    });

    $.each($('.soloNumero'), function (i, control) {
        $(control).attr("data-v-min", "0");
        $(control).attr("data-v-max", "99999999999");
        $(control).attr("data-m-dec", "0");
        $(control).attr("data-a-sep", "");
    });

    $.each($('.codigoPostal'), function (i, control) {
        $(control).attr("data-v-min", "0");
        $(control).attr("data-v-max", "9999999");
        $(control).attr("data-m-dec", "0");
        $(control).attr("data-a-sep", "");
    });

    $.each($('.soloOmdNumero'), function (i, control) {
        $(control).attr("data-v-min", "0");
        $(control).attr("data-v-max", "99999999999");
        $(control).attr("data-m-dec", "0");
        $(control).attr("data-a-sep", "");
    });

    $.each($('.soloOmdLinea'), function (i, control) {
        $(control).attr("data-v-min", "0");
        $(control).attr("data-v-max", "999999999999");
        $(control).attr("data-m-dec", "0");
        $(control).attr("data-a-sep", "");
    });

    $.each($('.montoFormato'), function (i, control) {
        $(control).attr("data-v-min", "0");
        $(control).attr("data-v-max", "99999999999999");
        $(control).attr("data-m-dec", "0");
        $(control).attr("data-a-sep", ".");
        $(control).attr("data-a-dec", ",");
    });

    $.each($('.dscr'), function (i, control) {
        $(control).attr("data-a-sep", ".");
        $(control).attr("data-a-dec", ",");
        $(control).attr("data-v-max", "999999999999.99");
        $(control).attr("data-m-dec", "2");
        $(control).attr("data-v-min", "0");
        $(control).attr("data-a-sign", "x");
        $(control).attr("data-p-sign", "s");

    });

    $('.rut').on('blur', function () {
        var rut = $(".rut").val();
        if (rutValido(rut)) {
            formateaRut(this, 'X.XXX.XXX-X');
        }
        else {
            alert("Rut Incorrecto");
            $(".rut").val("");
        }
    });

    $('.porcentajedosDecimales,.porcentajeUnDecimal,.porcentaje,.meses,.mesesDuration,.soloNumero,.dias,.soloOmdNumero,.soloOmdLinea').autoNumeric("init");
    $('.dosDecimales,.dosDecimalesTotal,.dscr,.montoFormato,.codigoPostal,.cuatroDecimales').autoNumeric("init");
    jQuery.validator.methods["date"] = function () { return true; };
});