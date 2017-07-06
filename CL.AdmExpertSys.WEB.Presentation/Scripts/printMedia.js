/**
 *  Version 2.4.0 Copyright (C) 2013
 *  Tested in IE 11, FF 28.0 and Chrome 33.0.1750.154
 *  No official support for other browsers, but will TRY to accommodate challenges in other browsers.
 *  Example:
 *      Print Button: <div id="print_button">Print</div>
 *      Print Area  : <div class="PrintArea" id="MyId" class="MyClass"> ... html ... </div>
 *      Javascript  : <script>
 *                       $("div#print_button").click(function(){
 *                           $("div.PrintArea").printArea( [OPTIONS] );
 *                       });
 *                     </script>
 *  options are passed as json (example: {mode: "popup", popClose: false})
 *
 *  {OPTIONS}   | [type]     | (default), values      | Explanation
 *  ---------   | ---------  | ---------------------- | -----------
 *  @mode       | [string]   | (iframe),popup         | printable window is either iframe or browser popup
 *  @popHt      | [number]   | (500)                  | popup window height
 *  @popWd      | [number]   | (400)                  | popup window width
 *  @popX       | [number]   | (500)                  | popup window screen X position
 *  @popY       | [number]   | (500)                  | popup window screen Y position
 *  @popTitle   | [string]   | ('')                   | popup window title element
 *  @popClose   | [boolean]  | (false),true           | popup window close after printing
 *  @extraCss   | [string]   | ('')                   | comma separated list of extra css to include
 *  @retainAttr | [string[]] | ["id","class","style"] | string array of attributes to retain for the containment area. (ie: id, style, class)
 *  @standard   | [string]   | strict, loose, (html5) | Only for popup. For html 4.01, strict or loose document standard, or html 5 standard
 *  @extraHead  | [string]   | ('')                   | comma separated list of extra elements to be appended to the head tag
 */

(function ($) {
    var counter = 0;
    var modes = { iframe: "iframe", popup: "popup" };
    var standards = { strict: "strict", loose: "loose", html5: "html5" };
    var defaults = {
        mode: modes.iframe,
        standard: standards.html5,
        popHt: 500,
        popWd: 400,
        popX: 200,
        popY: 200,
        popTitle: '',
        popClose: false,
        extraCss: '',
        extraHead: '',
        retainAttr: ["id", "class", "style"]
    };

    var settings = {};//global settings

    $.fn.printMedia = function (options) {
        $.extend(settings, defaults, options);

        counter++;
        var idPrefix = "printArea_";
        $("[id^=" + idPrefix + "]").remove();

        settings.id = idPrefix + counter;

        var $printSource = $(this);

        var printMediaWindow = PrintMedia.getPrintWindow();

        PrintMedia.write(printMediaWindow.doc, $printSource);

        setTimeout(function () { PrintMedia.print(printMediaWindow); }, 1000);
    };

    var PrintMedia = {
        print: function (PAWindow) {
            var paWindow = PAWindow.win;

            $(PAWindow.doc).ready(function () {
                paWindow.focus();
                paWindow.print();

                if (settings.mode == modes.popup && settings.popClose)
                    setTimeout(function () { paWindow.close(); }, 2000);
            });
        },
        write: function (PADocument, $ele) {
            PADocument.open();
            PADocument.write(PrintMedia.docTypeMedia() + "<html>" + PrintMedia.getHeadMedia() + PrintMedia.getBodyMedia($ele) + "</html>");
            PADocument.close();
        },
        docTypeMedia: function () {
            if (settings.mode == modes.iframe) return "";

            if (settings.standard == standards.html5) return "<!DOCTYPE html>";

            var transitional = settings.standard == standards.loose ? " Transitional" : "";
            var dtd = settings.standard == standards.loose ? "loose" : "strict";

            return '<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01' + transitional + '//EN" "http://www.w3.org/TR/html4/' + dtd + '.dtd">';
        },
        getHeadMedia: function () {
            var head = "<head><title>" + settings.popTitle + "</title>";
            $(document).find("link")
                .filter(function () {
                    return $(this).attr("rel").toLowerCase() == "stylesheet";
                })
                .filter(function () { // this filter contributed by "mindinquiring"
                    var media = $(this).attr("media");
                    return (media == "" || media == "print");
                })
                .each(function () {
                    head += '<link type="text/css" rel="stylesheet" href="' + $(this).attr("href") + '" >';
                });
            head += "</head>";
            return head;
        },
        getBodyMedia: function (printElement) {
            var titulo = $('.sub-header-titulo', $(printElement)).text();
            $(printElement).find('.sub-header-titulo').remove();

            var textoPlantilla = '';
            textoPlantilla += '<body>';
            textoPlantilla += ' <div align="center" border="0" cellpadding="0" cellspacing="0" width="100%">';
            textoPlantilla += '     <div class="cabezeraTitulo">';
            textoPlantilla += '             <div class="tituloImpresion">';
            textoPlantilla += '                 <img src="/Img/logo.png" style="width: 5.5em;"> <br /> <b> SANI Administración de negocio inmobiliario </b> ';
            textoPlantilla += '             </div>';
            textoPlantilla += '     </div>';
            if (titulo != '' && titulo != null)
            {
                textoPlantilla += '     <div>';
                textoPlantilla += '         <h2 class="sub-header">' + titulo + '</h2>';
                textoPlantilla += '     </div>';
            }
            textoPlantilla += '     <div>' + $(printElement).html() + '</div>';
            textoPlantilla += ' </div>';
            textoPlantilla += '</body>';
            return textoPlantilla;
        },
        getFormDataMedia: function (ele) {
            $("input,select,textarea", ele).each(function () {
                // In cases where radio, checkboxes and select elements are selected and deselected, and the print
                // button is pressed between select/deselect, the print screen shows incorrectly selected elements.
                // To ensure that the correct inputs are selected, when eventually printed, we must inspect each dom element
                var type = $(this).attr("type");
                if (type == "radio" || type == "checkbox") {
                    if ($(this).is(":not(:checked)")) this.removeAttribute("checked");
                    else this.setAttribute("checked", true);
                }
                else if (type == "text")
                    this.setAttribute("value", $(this).val());
                else if (type == "select-multiple" || type == "select-one" || $(this).get(0).tagName == 'SELECT')
                    $(this).find("option").each(function () {
                        if ($(this).is(":not(:selected)")) this.removeAttribute("selected");
                        else this.setAttribute("selected", true);
                    });
                else if (type == "textarea") {
                    var v = $(this).attr("value");
                    if ($.browser.mozilla) {
                        if (this.firstChild) this.firstChild.textContent = v;
                        else this.textContent = v;
                    }
                    else this.innerHTML = v;
                }
            });

            var elementoClone = $(ele).clone();

            $("input,select,textarea", elementoClone).each(function () {
                // In cases where radio, checkboxes and select elements are selected and deselected, and the print
                // button is pressed between select/deselect, the print screen shows incorrectly selected elements.
                // To ensure that the correct inputs are selected, when eventually printed, we must inspect each dom element
                var type = $(this).attr("type");
                if (type == "radio") {
                    if ($(this).is(":not(:checked)"))
                        this.removeAttribute("checked");
                    else
                        this.setAttribute("checked", true);
                }
                else if (type == "checkbox")
                {
                    if ($(this).is(":not(:checked)"))
                        $(this).replaceWith('<p>Si</p>');
                    else
                        $(this).replaceWith('<p>No</p>');

                }
                else if (type == "text" || type == "numeric")
                    $(this).replaceWith('<p>' + $(this).val() + '</p>');
                else if ($(this).get(0).tagName == 'SELECT') {
                    if ($(this).val() != '' && $(this).val() != '0' && $(this).val() != null && $(this).val() != '--Seleccione--' && $(this).val() != '') {
                        $(this).replaceWith('<p>' + $('option:selected', $(this)).text() + '</p>');
                    }
                    else
                        $(this).remove();
                }
                else if ($(this).get(0).tagName == 'TEXTAREA') {
                    $(this).replaceWith('<p>' + $(this).val() + '</p>');
                }
            });


            return elementoClone;
        },
        getPrintWindow: function () {
            switch (settings.mode) {
                case modes.iframe:
                    var f = new PrintMedia.IframeMedia();
                    return { win: f.contentWindow || f, doc: f.doc };
                case modes.popup:
                    var p = new PrintMedia.PopupMedia();
                    return { win: p, doc: p.doc };
            }
        },
        IframeMedia: function () {
            var frameId = settings.id;
            var iframeStyle = 'border:0;visibility:hidden;';
            var iframe;

            try {
                iframe = document.createElement('iframe');
                document.body.appendChild(iframe);
                $(iframe).attr({ style: iframeStyle, id: frameId, src: "" });
                iframe.doc = null;
                iframe.doc = iframe.contentDocument ? iframe.contentDocument : (iframe.contentWindow ? iframe.contentWindow.document : iframe.document);
            }
            catch (e) { throw e + ". iframes may not be supported in this browser."; }

            if (iframe.doc == null) throw "Cannot find document.";

            return iframe;
        },
        PopupMedia: function () {
            var windowAttr = "location=yes,statusbar=no,directories=no,menubar=no,titlebar=no,toolbar=no,dependent=no";
            windowAttr += ",width=" + settings.popWd + ",height=" + settings.popHt;
            windowAttr += ",resizable=yes,screenX=" + settings.popX + ",screenY=" + settings.popY + ",personalbar=no,scrollbars=no";

            var newWin = window.open("", "_blank", windowAttr);

            newWin.doc = newWin.document;

            return newWin;
        }
    };
})(jQuery);