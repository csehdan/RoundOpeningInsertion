using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;

namespace RoundOpeningInsertion
{
	[Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]

    public class RoundOpeningInsertion : IExternalCommand
    {
        private IAutoCreateObjects autoCreateObjects;

        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            autoCreateObjects = new RoundOpeningCreator(commandData); // I have not found the right way to use DI for this runtime parameter
            return autoCreateObjects.AutoCreateObjects();
        }
    }
}
