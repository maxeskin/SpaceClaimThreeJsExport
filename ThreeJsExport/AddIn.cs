using System;
using System.IO;
using System.Windows.Forms;
using System.Collections.Generic;
using System.Diagnostics;
using System.Text;
using System.Xml;
using System.Linq;
using SpaceClaim.Api.V11;
using SpaceClaim.Api.V11.Extensibility;
using SpaceClaim.Api.V11.Geometry;
using SpaceClaim.Api.V11.Modeler;
using SpaceClaim.Api.V11.Display;
using ThreeJsExport.Properties;
using Color = System.Drawing.Color;
using Application = SpaceClaim.Api.V11.Application;
using System.Drawing;

namespace SpaceClaim.AddIn.ThreeJsExport {
	public class AddIn : SpaceClaim.Api.V11.Extensibility.AddIn, IExtensibility, IRibbonExtensibility, ICommandExtensibility {
		public bool Connect() {
			return true;
		}

		public void Disconnect() {
		}

		public string GetCustomUI() {
            string test = Resources.Ribbon;
			return test;
		}

        CommandCapsule button;

		public void Initialize() {
            button = new ExportThreeJsToolCapsule("ThreeJsExport");
            button.Initialize();

            var surfaceDeviationCommand = Command.Create("SurfaceDeviation");
            surfaceDeviationCommand.Text = (new TessellationOptions()).SurfaceDeviation.ToString();

            var angleDeviationCommand = Command.Create("AngleDeviation");
            angleDeviationCommand.Text = (new TessellationOptions()).AngleDeviation.ToString();

        }
	}
}
