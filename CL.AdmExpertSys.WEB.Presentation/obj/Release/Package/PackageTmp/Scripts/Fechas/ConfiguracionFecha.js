$('input.fecha').after("<span class='input-group-addon'><i class='glyphicon glyphicon-calendar'></i></span>");
$("input.fecha").attr("readonly", "readonly");
$("input.fecha").attr("style", "cursor:pointer");

$.each($('.fechaContainer .input-group.date'), function (i, control) {
    $(control).datepicker({        
        format: "dd-mm-yyyy",
        minViewMode: 3,
        autoclose: true,
        language: 'es',
        orientation: "top left"
    });
});