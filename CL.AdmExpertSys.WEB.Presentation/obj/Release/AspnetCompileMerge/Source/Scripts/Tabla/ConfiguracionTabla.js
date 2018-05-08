$(document).ready(function () {
    $('.table.table-hover.tablaMantenedores').dataTable({
        "dom": '<"pull-left"f><"pull-right"T><"pull-right"l>tip',
        "lengthMenu": [[10, 30, 50, 100], [10, 30, 50, 100]],
        "order": [],
        "tableTools": {
            "sSwfPath": '../Scripts/TableTools/swf/copy_csv_xls_pdf.swf',            
            "aButtons": [
                //"print",
                {
                    "sExtends": "collection",
                    "aButtons": [
                    {
                         "sExtends": "xls",
                         "mColumns": [0, 1, 2, 3]
                    },
                    {
                        "sExtends": "csv",
                        "mColumns": [0, 1, 2, 3]
                    },
                    {
                        "sExtends": "pdf",
                        "mColumns": [0, 1, 2, 3]
                    }],
                }
            ]
        },
        "language": {
            "sProcessing": "Procesando...",
            "sLengthMenu": "Mostrar _MENU_" + "&nbsp;&nbsp;&nbsp;",
            "sZeroRecords": "No se encontraron resultados",
            "sEmptyTable": "Ningún dato disponible en esta tabla",
            "sInfo": "Mostrando registros del _START_ al _END_ de un total de _TOTAL_ registros",
            "sInfoEmpty": "Mostrando registros del 0 al 0 de un total de 0 registros",
            "sInfoFiltered": "(filtrado de un total de _MAX_ registros)",
            "sInfoPostFix": "",
            "sSearch": "Buscar:",
            "sUrl": "",
            "sInfoThousands": ",",
            "sLoadingRecords": "Cargando...",
            "oPaginate": {
                "sFirst": "Primero",
                "sLast": "Último",
                "sNext": "Siguiente",
                "sPrevious": "Anterior"
            },
            "oAria": {
                "sSortAscending": ": Activar para ordenar la columna de manera ascendente",
                "sSortDescending": ": Activar para ordenar la columna de manera descendente"
            }
        }
    });
});