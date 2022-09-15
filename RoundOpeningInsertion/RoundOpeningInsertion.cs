using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;

namespace RoundOpeningInsertion
{
	[Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]

    public class RoundOpeningInsertion : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            var roundOpeningCreator = new RoundOpeningCreator(commandData);
            return roundOpeningCreator.CreateRoundOpenings();
        }
    }
}
