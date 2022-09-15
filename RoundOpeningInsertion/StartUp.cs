using Microsoft.Extensions.DependencyInjection;

namespace RoundOpeningInsertion
{
	public class StartUp
	{
		public void ConfigureServices(IServiceCollection serviceCollection)
		{
			serviceCollection.AddScoped<IAutoCreateObjects, RoundOpeningCreator>();
		}
	}
}
