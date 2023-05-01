using System;
using System.Collections.Generic;
using System.Text;
using System.IO;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using ESRI.ArcGIS.ArcMapUI;
using ESRI.ArcGIS.Carto;
using ESRI.ArcGIS.esriSystem;
using ESRI.ArcGIS.Geodatabase;
using Serilog;

namespace changeVersionAllLayers
{
    public class changevesion : ESRI.ArcGIS.Desktop.AddIns.Button
    {
        public static IWorkspace _currentWorkspace;

        public changevesion()
        {
            try
            {
                IContentsView pTOC = ArcMap.Document.CurrentContentsView;
                ILayer layer = (ILayer)pTOC.SelectedItem;
                IFeatureLayer featurelayer = (IFeatureLayer)layer;
                var database = (IDataset)featurelayer;
                _currentWorkspace = database.Workspace;
            }
            catch (Exception e)
            {
                MessageBox.Show("Sélectionner la couche à Identifier");
                
                return; 
            }
            
        }

        protected override void OnClick()

        {
            Log.Logger = new LoggerConfiguration()
                .WriteTo.Console()
                .WriteTo.File(@"C:\Users\lf.collazos\source\repos\changeVersionAllLayers\logs\log.txt")
                .CreateLogger();
            Log.Information("Start JOB");
            
            var server = _currentWorkspace.ConnectionProperties.GetProperty("SERVER");
            var instance = _currentWorkspace.ConnectionProperties.GetProperty("INSTANCE");
            var user = _currentWorkspace.ConnectionProperties.GetProperty("USER");
            var password = _currentWorkspace.ConnectionProperties.GetProperty("PASSWORD");
            var version = _currentWorkspace.ConnectionProperties.GetProperty("VERSION");

            Log.Information((string)instance);
            var pEnumLayer = ArcMap.Document.ActivatedView.FocusMap.Layers;
            var pLayer = pEnumLayer.Next();
            do
            {
                if (pLayer == null) break;
                if (pLayer is FeatureLayer)
                {
                    ChangeWorkspaceSde(pLayer, server, instance, user, password, version);
                }

                pLayer = pEnumLayer.Next();
            } while (true);
            Log.Information("End JOB");
            Log.CloseAndFlush();
           
        }

        public static void ChangeWorkspaceSde(ILayer layer, object server, object instance, object user, object password, object version)
        {
            try
            {

                var dataLayer = (IDataLayer2)layer;
            var name = dataLayer.DataSourceName;
            var datasetName = (IDatasetName)dataLayer.DataSourceName;
            var workspaceName = datasetName.WorkspaceName;
            dataLayer.Disconnect();
            Log.Information("changenment instance= " + layer.Name);
            Log.Information((string)workspaceName.ConnectionProperties.GetProperty("INSTANCE"));

            if ((string) workspaceName.ConnectionProperties.GetProperty("INSTANCE") == "5152")
            {
                
                var propertySet = new PropertySet();

                {
                    propertySet.SetProperty("SERVER", server);
                    propertySet.SetProperty("INSTANCE", instance);
                    propertySet.SetProperty("USER", user);
                    propertySet.SetProperty("PASSWORD", password);
                    propertySet.SetProperty("VERSION", version);
                    workspaceName.ConnectionProperties = propertySet;
                }
            }
 
                dataLayer.Connect(name);
                ArcMap.Document.UpdateContents();

                Marshal.ReleaseComObject(_currentWorkspace);
                Marshal.ReleaseComObject(workspaceName);
                _currentWorkspace = null;
                GC.Collect();
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message);
                MessageBox.Show(layer.Name);
            }

            

        }
      
        protected override void OnUpdate()
        {
            Enabled = ArcMap.Application != null;
        }
    }

}
