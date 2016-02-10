using System.Web.Http;
using WebActivatorEx;
using ExporterWeb;
using Swashbuckle.MVC.Handler;
using Microsoft.Web.Infrastructure.DynamicModuleHelper;

[assembly: PreApplicationStartMethod(typeof(SwaggerMVCConfig), "Register")]
namespace ExporterWeb
{
    public class SwaggerMVCConfig
    {
		public static void Register()
        {
			DynamicModuleUtility.RegisterModule(typeof(SwashbuckleMVCModule));
		}
	}
}