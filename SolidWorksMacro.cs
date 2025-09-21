using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Forms;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;


namespace Macro2
{
    public partial class SolidWorksMacro
    {
        public SldWorks swApp;
        private const string BASE_LENGTH = "150mm";
        private const string BASE_HEIGHT = "140mm";
        public void Main()
        {
            swApp = (SldWorks)Marshal.GetActiveObject("SldWorks.Application");
            swApp.SendMsgToUser("Macro started successfully.");

            // Optional: reduces other UI popups while running API code
            bool prevCmdInProgress = swApp.CommandInProgress;
            swApp.CommandInProgress = true;


            AddOrUpdateGlobalVariable("BaseLength", BASE_LENGTH);
            AddOrUpdateGlobalVariable("BaseHeight", BASE_HEIGHT);
            AddOrUpdateConfiguration("pla");
            CreateCenterRectangleSketch();

            swApp.SendMsgToUser("Macro completed successfully.");
            return;
        }

        /// <summary>
        /// Adds or updates a global variable in the active model.
        /// </summary>
        /// <param name="name">Name of the global variable</param>
        /// <param name="value">Value expression (e.g. "100", "50mm", "Width*2")</param>
        private void AddOrUpdateGlobalVariable(string name, string value)
        {
            ModelDoc2 model = swApp.ActiveDoc as ModelDoc2;
            if (model == null)
            {
                swApp.SendMsgToUser("No active document.");
                return;
            }

            EquationMgr eqMgr = model.GetEquationMgr();
            int count = eqMgr.GetCount();

            // Look for existing variable
            int existingIndex = -1;
            for (int i = 0; i < count; i++)
            {
                string eq = eqMgr.Equation[i];
                // Global variables are stored like `"\"VarName\" = Value"`
                if (eq.StartsWith("\"" + name + "\""))
                {

                    existingIndex = i;
                    break;
                }
            }

            string equationString = "\"" + name + "\"=" + value;

            if (existingIndex >= 0)
            {
                // Update existing
                eqMgr.Equation[existingIndex] = equationString;
            }
            else
            {
                // Add new
                eqMgr.Add2(count, equationString, true);
            }

            // Force rebuild to apply changes
            model.EditRebuild3();
        }

        private void AddOrUpdateConfiguration(string configName, string comment = "", string altName = "")
        {
            ModelDoc2 model = swApp.ActiveDoc as ModelDoc2;
            if (model == null)
            {
                swApp.SendMsgToUser("No active document.");
                return;
            }

            ConfigurationManager configMgr = model.ConfigurationManager;
            if (configMgr == null)
            {
                swApp.SendMsgToUser("No configuration manager available.");
                return;
            }

            // See if configuration already exists
            Configuration existingConfig = (Configuration)model.GetConfigurationByName(configName);

            if (existingConfig == null)
            {
                // Create new configuration
                // Parameters: Name, Options, Comment, Alternate name
                // Options can be swConfigurationOptions2_e enum, e.g. swConfigOption_DontActivate
                configMgr.AddConfiguration(configName,
                    comment,
                    altName,
                    (int)swConfigurationOptions2_e.swConfigOption_DontActivate,
                    "",
                    "");

                swApp.SendMsgToUser("Configuration '" + configName + "' created.");
            }
            else
            {
                // Update comment or alternate name
                existingConfig.Description = comment;
                existingConfig.AlternateName = altName;

                swApp.SendMsgToUser("Configuration '" + configName + "' updated.");
            }
        }

        /// <summary>
        /// Creates a new sketch with a center point rectangle dimensioned to BaseLength and BaseHeight
        /// </summary>
        private void CreateCenterRectangleSketch()
        {
            ModelDoc2 model = swApp.ActiveDoc as ModelDoc2;
            if (model == null)
            {
                swApp.SendMsgToUser("No active document.");
                return;
            }

            // Check if this is a part document
            if (model.GetType() != (int)swDocumentTypes_e.swDocPART)
            {
                swApp.SendMsgToUser("Active document must be a part.");
                return;
            }

            PartDoc part = (PartDoc)model;

            // Select the Front Plane (or another plane as needed)
            bool boolstatus = model.Extension.SelectByID2("Front Plane", "PLANE", 0, 0, 0, false, 0, null, 0);

            if (!boolstatus)
            {
                // Try selecting "Front" if "Front Plane" doesn't work
                boolstatus = model.Extension.SelectByID2("Front", "PLANE", 0, 0, 0, false, 0, null, 0);

                if (!boolstatus)
                {
                    swApp.SendMsgToUser("Could not select Front Plane. Please select a plane manually and re-run the macro.");
                    return;
                }
            }

            model.SketchManager.InsertSketch(true);

            double halfSize = 0.075;

            // don't make it a square unless that's what you want
            object rect = model.SketchManager.CreateCenterRectangle(0, 0, 0, halfSize, .080, 0);

            model.ClearSelection2(true);

            // Get all sketch segments
            object[] sketchSegments = (object[])model.SketchManager.ActiveSketch.GetSketchSegments();

            if (sketchSegments != null && sketchSegments.Length >= 4)
            {
                // A rectangle has 4 lines
                // Select first line (should be horizontal) for width
                SketchSegment line1 = (SketchSegment)sketchSegments[0];
                if (!line1.Select4(false, null))
                {
                    swApp.SendMsgToUser("Could not select line1");
                    return;
                }

                // Add horizontal dimension
                DisplayDimension dimWidth = model.AddDimension2(0, halfSize * 1.5, 0) as DisplayDimension;

                if (dimWidth != null)
                {
                    Dimension dim = (Dimension)dimWidth.GetDimension();
                    string dimName = dim.FullName;
                    swApp.SendMsgToUser($"dim {dim.Name} @Sketch1");
                    string equation = "\"" + dimName + "\" = \"BaseLength\"";
                    model.GetEquationMgr().Add2(-1, equation, true);
                }

                model.ClearSelection2(true);

                // Select second line (should be vertical) for height
                SketchSegment line2 = (SketchSegment)sketchSegments[1];
                if (!line2.Select4(false, null))
                {
                    swApp.SendMsgToUser("Could not select line2");
                    return;
                }

                // Add vertical dimension
                DisplayDimension dimHeight = model.AddDimension2(halfSize * 1.5, 0, 0) as DisplayDimension;

                if (dimHeight != null)
                {
                    Dimension dim = (Dimension)dimHeight.GetDimension();
                    swApp.SendMsgToUser(dim.FullName);
                    string dimName = dim.FullName;
                    string equation = "\"" + dimName + "\" = \"BaseHeight\"";
                    model.GetEquationMgr().Add2(-1, equation, true);
                }
            }
            else
            {
                swApp.SendMsgToUser("Could not get sketch segments for dimensioning.");
            }

            model.ClearSelection2(true);

            // Exit sketch mode
            model.SketchManager.InsertSketch(true);

            // Rebuild to apply all changes
            model.EditRebuild3();
        }
    }
}

