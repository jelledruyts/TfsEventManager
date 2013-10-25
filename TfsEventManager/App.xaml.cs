using System.Globalization;
using System.Windows;
using System.Windows.Markup;

namespace TfsEventManager
{
    public partial class App : Application
    {
        protected override void OnStartup(StartupEventArgs e)
        {
            // Ensure that the current culture is used for all controls (see http://www.west-wind.com/Weblog/posts/796725.aspx).
            FrameworkElement.LanguageProperty.OverrideMetadata(typeof(FrameworkElement), new FrameworkPropertyMetadata(XmlLanguage.GetLanguage(CultureInfo.CurrentCulture.IetfLanguageTag)));
        }
    }
}