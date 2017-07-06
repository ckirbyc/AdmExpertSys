
using System.Collections.Generic;
using System.Linq;
using System.Web.Mvc;

namespace CL.AdmExpertSys.WEB.Presentation.Mapping.Factories
{
    public class BaseFactory
    {
        public List<SelectListItem> GetDropDownList<T>(IList<T> items, string text, string value, string selected) where T : class
        {
            var list = (from item in items
                        select item).AsEnumerable().Select(m => new SelectListItem
                        {
                            Text = m.GetType().GetProperty(text).GetValue(m).ToString(),
                            Value = m.GetType().GetProperty(value).GetValue(m).ToString(),
                            Selected = selected != string.Empty &&
                                 ReferenceEquals(m.GetType().GetProperty(value).GetValue(m), selected),
                        }).ToList();
            return list;
        }
    }
}
