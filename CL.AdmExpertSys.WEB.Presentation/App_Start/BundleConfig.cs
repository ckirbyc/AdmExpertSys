using System.Web.Optimization;

namespace CL.AdmExpertSys.WEB.Presentation
{
    public class BundleConfig
    {
        // Para obtener más información acerca de Bundling, consulte http://go.microsoft.com/fwlink/?LinkId=254725
        public static void RegisterBundles(BundleCollection bundles)
        {
            bundles.Add(new ScriptBundle("~/bundles/jquery").Include(
                "~/Scripts/jquery-{version}.js",
                "~/Scripts/jquery.validate.js",
                "~/Scripts/Select2/select2.full.js",
                "~/Scripts/jquery.dataTables.js",
                "~/Scripts/jquery.dataTables.rowGrouping.js",
                "~/Scripts/TableTools/js/dataTables.tableTools.js",
                "~/Scripts/jquery.base64.js",
                "~/Scripts/jstree/jstree.min.js",
                "~/Scripts/jquery.cookie.js"
                ));

            bundles.Add(new ScriptBundle("~/bundles/bootstrap").Include(
                    "~/Scripts/bootstrap.js",
                    "~/Scripts/bootstrap-dialog.min.js",
                    "~/Scripts/bootstrap3-confirmation.js",
                    "~/Scripts/bootstrap-table.min.js",
                    "~/Scripts/dataTables.bootstrap.js",
                    "~/Scripts/bootstrap-datepicker.js",
                    "~/Scripts/bootstrap-datepicker.es.js",
                    "~/Scripts/jasny-bootstrap.min.js",
                    "~/Scripts/moment.min.js",
                    "~/Scripts/datatime-moment.js",
                    "~/Scripts/printArea.js",
                    "~/Scripts/printMedia.js"
                ));

            bundles.Add(new ScriptBundle("~/bundles/otros").Include(
                    "~/Scripts/tableExport.js",
                    "~/Scripts/autoNumeric.js",
                    "~/Scripts/main.js"
                ));

            bundles.Add(new StyleBundle("~/Content/css").Include(
                "~/Css/bootstrap.css",
                //"~/Css/bootstrap-dialog.min.css",
                "~/Css/Select2/select2.css",
                "~/Css/bootstrap-table.min.css",
                "~/Css/dashboard.css",
                "~/Css/dataTables.bootstrap.css",
                "~/Css/datepicker3.css",
                "~/Css/jasny-bootstrap.min.css",
                "~/Css/font-awesome.min.css",
                "~/Scripts/TableTools/css/dataTables.tableTools.css",
                "~/Css/bootstrap-min.css",
                "~/Css/tree/style.min.css"
                ));
            BundleTable.EnableOptimizations = true;
        }
    }
}