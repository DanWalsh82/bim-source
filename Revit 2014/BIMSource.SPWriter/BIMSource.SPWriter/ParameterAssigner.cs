using Autodesk.Revit.DB;
using System;
using System.IO;
using System.Windows.Forms;

namespace BIMSource.SPWriter
{
  class ParameterAssigner
  {
    private Autodesk.Revit.ApplicationServices.Application m_app;
    private FamilyManager m_manager = null;

    private DefinitionFile m_sharedFile;
    private string m_sharedFilePath = string.Empty;

    public ParameterAssigner(Autodesk.Revit.ApplicationServices.Application app, Document doc)
    {
        m_app = app;
        m_manager = doc.FamilyManager;
    }

    public bool LoadSharedParameterFile()
    {
      string myDocsFolder = Environment.GetFolderPath(Environment.SpecialFolder.Personal);
      OpenFileDialog ofd = new OpenFileDialog();

      ofd.InitialDirectory = myDocsFolder;
      ofd.Filter = "txt files (*.txt)|*.txt|All files (*.*)|*.*";
      ofd.FilterIndex = 1;
      ofd.RestoreDirectory = true;
      ofd.Title = "Please Select the Shared Parameter File";

      if (ofd.ShowDialog() == DialogResult.OK)
      {
        m_sharedFilePath = ofd.FileName;
        if (!File.Exists(m_sharedFilePath))
        {
          return true;
        }

        m_app.SharedParametersFilename = m_sharedFilePath;
        try
        {
          m_sharedFile = m_app.OpenSharedParameterFile();
        }
        catch (System.Exception e)
        {
          MessageBox.Show(e.Message);
          return false;
        }
        return true;
      }
      else
      {
        return false;
      }
    }

    public bool BindSharedParameters()
    {
      if (File.Exists(m_sharedFilePath) &&
          null == m_sharedFile)
      {
        MessageBox.Show("SharedParameter.txt has an invalid format.");
        return false;
      }

      foreach (DefinitionGroup group in m_sharedFile.Groups)
      {
        foreach (ExternalDefinition def in group.Definitions)
        {
          // check whether the parameter already exists in the document
          FamilyParameter param = m_manager.get_Parameter(def.Name);
          if (null != param)
          {
            continue;
          }

          BuiltInParameterGroup bpg = BuiltInParameterGroup.INVALID;
          try
          {
            if (def.OwnerGroup.Name == "Dimensions")
            {
              bpg = BuiltInParameterGroup.PG_GEOMETRY;
            }
            else if (def.OwnerGroup.Name == "ADSK Model Properties")
            {
              bpg = BuiltInParameterGroup.PG_ADSK_MODEL_PROPERTIES  ;
            }
            else if (def.OwnerGroup.Name == "Aelectrical")
            {
              bpg = BuiltInParameterGroup.PG_AELECTRICAL  ;
            }
            else if (def.OwnerGroup.Name == "Analysis Results")
            {
              bpg = BuiltInParameterGroup.PG_ANALYSIS_RESULTS  ;
            }
            else if (def.OwnerGroup.Name == "Analytical Alignment")
            {
              bpg = BuiltInParameterGroup.PG_ANALYTICAL_ALIGNMENT  ;
            }
            else if (def.OwnerGroup.Name == "Analytical Model")
            {
              bpg = BuiltInParameterGroup.PG_ANALYTICAL_MODEL  ;
            }
            else if (def.OwnerGroup.Name == "Analytical Properties")
            {
              bpg = BuiltInParameterGroup.PG_ANALYTICAL_PROPERTIES  ;
            }
            else if (def.OwnerGroup.Name == "Area")
            {
              bpg = BuiltInParameterGroup.PG_AREA  ;
            }
            else if (def.OwnerGroup.Name == "Conceptual Energy Data")
            {
              bpg = BuiltInParameterGroup.PG_CONCEPTUAL_ENERGY_DATA  ;
            }
            else if (def.OwnerGroup.Name == "Conceptual Energy Data Building Services")
            {
              bpg = BuiltInParameterGroup.PG_CONCEPTUAL_ENERGY_DATA_BUILDING_SERVICES  ;
            }
            else if (def.OwnerGroup.Name == "Constraints")
            {
              bpg = BuiltInParameterGroup.PG_CONSTRAINTS  ;
            }
            else if (def.OwnerGroup.Name == "Construction")
            {
              bpg = BuiltInParameterGroup.PG_CONSTRUCTION  ;
            }
            else if (def.OwnerGroup.Name == "Continuousrail Begin Bottom Extension")
            {
              bpg = BuiltInParameterGroup.PG_CONTINUOUSRAIL_BEGIN_BOTTOM_EXTENSION  ;
            }
            else if (def.OwnerGroup.Name == "Continuousrail End Top Extension")
            {
              bpg = BuiltInParameterGroup.PG_CONTINUOUSRAIL_END_TOP_EXTENSION  ;
            }
            else if (def.OwnerGroup.Name == "Data")
            {
              bpg = BuiltInParameterGroup.PG_DATA  ;
            }
            else if (def.OwnerGroup.Name == "Display")
            {
              bpg = BuiltInParameterGroup.PG_DISPLAY  ;
            }
            else if (def.OwnerGroup.Name == "Electrical")
            {
              bpg = BuiltInParameterGroup.PG_ELECTRICAL  ;
            }
            else if (def.OwnerGroup.Name == "Electrical - Loads")
            {
              bpg = BuiltInParameterGroup.PG_ELECTRICAL_LOADS  ;
            }
            else if (def.OwnerGroup.Name == "Electrical Circuiting")
            {
              bpg = BuiltInParameterGroup.PG_ELECTRICAL_CIRCUITING  ;
            }
            else if (def.OwnerGroup.Name == "Electrical Lighting")
            {
              bpg = BuiltInParameterGroup.PG_ELECTRICAL_LIGHTING  ;
            }
            else if (def.OwnerGroup.Name == "Energy Analysis")
            {
              bpg = BuiltInParameterGroup.PG_ENERGY_ANALYSIS  ;
            }
            else if (def.OwnerGroup.Name == "Energy Analysis Conceptual Model")
            {
              bpg = BuiltInParameterGroup.PG_ENERGY_ANALYSIS_CONCEPTUAL_MODEL  ;
            }
            else if (def.OwnerGroup.Name == "Energy Analysis Detailed And Conceptual Models")
            {
              bpg = BuiltInParameterGroup.PG_ENERGY_ANALYSIS_DETAILED_AND_CONCEPTUAL_MODELS  ;
            }
            else if (def.OwnerGroup.Name == "Energy Analysis Detailed Model")
            {
              bpg = BuiltInParameterGroup.PG_ENERGY_ANALYSIS_DETAILED_MODEL  ;
            }
            else if (def.OwnerGroup.Name == "Fire Protection")
            {
              bpg = BuiltInParameterGroup.PG_FIRE_PROTECTION  ;
            }
            else if (def.OwnerGroup.Name == "Fitting")
            {
              bpg = BuiltInParameterGroup.PG_FITTING  ;
            }
            else if (def.OwnerGroup.Name == "Flexible")
            {
              bpg = BuiltInParameterGroup.PG_FLEXIBLE  ;
            }
            else if (def.OwnerGroup.Name == "General")
            {
              bpg = BuiltInParameterGroup.PG_GENERAL  ;
            }
            else if (def.OwnerGroup.Name == "Graphics")
            {
              bpg = BuiltInParameterGroup.PG_GRAPHICS  ;
            }
            else if (def.OwnerGroup.Name == "Green Building Properties")
            {
              bpg = BuiltInParameterGroup.PG_GREEN_BUILDING  ;
            }
            else if (def.OwnerGroup.Name == "Identity Data")
            {
              bpg = BuiltInParameterGroup.PG_IDENTITY_DATA  ;
            }
            else if (def.OwnerGroup.Name == "IFC")
            {
              bpg = BuiltInParameterGroup.PG_IFC  ;
            }
            else if (def.OwnerGroup.Name == "Insulation")
            {
              bpg = BuiltInParameterGroup.PG_INSULATION  ;
            }
            else if (def.OwnerGroup.Name == "Length")
            {
              bpg = BuiltInParameterGroup.PG_LENGTH  ;
            }
            else if (def.OwnerGroup.Name == "Light Photometrics")
            {
              bpg = BuiltInParameterGroup.PG_LIGHT_PHOTOMETRICS  ;
            }
            else if (def.OwnerGroup.Name == "Lining")
            {
              bpg = BuiltInParameterGroup.PG_LINING  ;
            }
            else if (def.OwnerGroup.Name == "Materials and Finishes")
            {
              bpg = BuiltInParameterGroup.PG_MATERIALS  ;
            }
            else if (def.OwnerGroup.Name == "Mechanical")
            {
              bpg = BuiltInParameterGroup.PG_MECHANICAL  ;
            }
            else if (def.OwnerGroup.Name == "Mechanical - Air Flow")
            {
              bpg = BuiltInParameterGroup.PG_MECHANICAL_AIRFLOW  ;
            }
            else if (def.OwnerGroup.Name == "Mechanical - Loads")
            {
              bpg = BuiltInParameterGroup.PG_MECHANICAL_LOADS  ;
            }
            else if (def.OwnerGroup.Name == "Nodes")
            {
              bpg = BuiltInParameterGroup.PG_NODES  ;
            }
            else if (def.OwnerGroup.Name == "Orientation")
            {
              bpg = BuiltInParameterGroup.PG_ORIENTATION  ;
            }
            else if (def.OwnerGroup.Name == "Overall Legend")
            {
              bpg = BuiltInParameterGroup.PG_OVERALL_LEGEND  ;
            }
            else if (def.OwnerGroup.Name == "Pattern")
            {
              bpg = BuiltInParameterGroup.PG_PATTERN  ;
            }
            else if (def.OwnerGroup.Name == "Pattern Application")
            {
              bpg = BuiltInParameterGroup.PG_PATTERN_APPLICATION  ;
            }
            else if (def.OwnerGroup.Name == "Pattern Grid")
            {
              bpg = BuiltInParameterGroup.PG_PATTERN_GRID  ;
            }
            else if (def.OwnerGroup.Name == "Pattern Grid 1")
            {
              bpg = BuiltInParameterGroup.PG_PATTERN_GRID_1  ;
            }
            else if (def.OwnerGroup.Name == "Pattern Grid 2")
            {
              bpg = BuiltInParameterGroup.PG_PATTERN_GRID_2  ;
            }
            else if (def.OwnerGroup.Name == "Pattern Grid Horiz")
            {
              bpg = BuiltInParameterGroup.PG_PATTERN_GRID_HORIZ  ;
            }
            else if (def.OwnerGroup.Name == "Pattern Grid U")
            {
              bpg = BuiltInParameterGroup.PG_PATTERN_GRID_U  ;
            }
            else if (def.OwnerGroup.Name == "Pattern Grid V")
            {
              bpg = BuiltInParameterGroup.PG_PATTERN_GRID_V  ;
            }
            else if (def.OwnerGroup.Name == "Pattern Grid Vert")
            {
              bpg = BuiltInParameterGroup.PG_PATTERN_GRID_VERT  ;
            }
            else if (def.OwnerGroup.Name == "Pattern Mullion 1")
            {
              bpg = BuiltInParameterGroup.PG_PATTERN_MULLION_1  ;
            }
            else if (def.OwnerGroup.Name == "Pattern Mullion 2")
            {
              bpg = BuiltInParameterGroup.PG_PATTERN_MULLION_2  ;
            }
            else if (def.OwnerGroup.Name == "Pattern Mullion Horiz")
            {
              bpg = BuiltInParameterGroup.PG_PATTERN_MULLION_HORIZ  ;
            }
            else if (def.OwnerGroup.Name == "Pattern Mullion Vert")
            {
              bpg = BuiltInParameterGroup.PG_PATTERN_MULLION_VERT  ;
            }
            else if (def.OwnerGroup.Name == "Phasing")
            {
              bpg = BuiltInParameterGroup.PG_PHASING  ;
            }
            else if (def.OwnerGroup.Name == "Plumbing")
            {
              bpg = BuiltInParameterGroup.PG_PLUMBING  ;
            }
            else if (def.OwnerGroup.Name == "Profile")
            {
              bpg = BuiltInParameterGroup.PG_PROFILE  ;
            }
            else if (def.OwnerGroup.Name == "Profile 1")
            {
              bpg = BuiltInParameterGroup.PG_PROFILE_1  ;
            }
            else if (def.OwnerGroup.Name == "Profile 2")
            {
              bpg = BuiltInParameterGroup.PG_PROFILE_2  ;
            }
            else if (def.OwnerGroup.Name == "Railing System Family Handrails")
            {
              bpg = BuiltInParameterGroup.PG_RAILING_SYSTEM_FAMILY_HANDRAILS  ;
            }
            else if (def.OwnerGroup.Name == "Railing System Family Segment Pattern")
            {
              bpg = BuiltInParameterGroup.PG_RAILING_SYSTEM_FAMILY_SEGMENT_PATTERN  ;
            }
            else if (def.OwnerGroup.Name == "Railing System Family Top Rail")
            {
              bpg = BuiltInParameterGroup.PG_RAILING_SYSTEM_FAMILY_TOP_RAIL  ;
            }
            else if (def.OwnerGroup.Name == "Railing System Secondary Family Handrails")
            {
              bpg = BuiltInParameterGroup.PG_RAILING_SYSTEM_SECONDARY_FAMILY_HANDRAILS  ;
            }
            else if (def.OwnerGroup.Name == "Railing System Segment Pattern Remainder")
            {
              bpg = BuiltInParameterGroup.PG_RAILING_SYSTEM_SEGMENT_PATTERN_REMAINDER  ;
            }
            else if (def.OwnerGroup.Name == "Railing System Segment Pattern Repeat")
            {
              bpg = BuiltInParameterGroup.PG_RAILING_SYSTEM_SEGMENT_PATTERN_REPEAT  ;
            }
            else if (def.OwnerGroup.Name == "Railing System Segment Posts")
            {
              bpg = BuiltInParameterGroup.PG_RAILING_SYSTEM_SEGMENT_POSTS  ;
            }
            else if (def.OwnerGroup.Name == "Railing System Segment U Grid")
            {
              bpg = BuiltInParameterGroup.PG_RAILING_SYSTEM_SEGMENT_U_GRID  ;
            }
            else if (def.OwnerGroup.Name == "Railing System Segment V Grid")
            {
              bpg = BuiltInParameterGroup.PG_RAILING_SYSTEM_SEGMENT_V_GRID  ;
            }
            else if (def.OwnerGroup.Name == "Rebar Array")
            {
              bpg = BuiltInParameterGroup.PG_REBAR_ARRAY  ;
            }
            else if (def.OwnerGroup.Name == "Rebar System Layers")
            {
              bpg = BuiltInParameterGroup.PG_REBAR_SYSTEM_LAYERS  ;
            }
            else if (def.OwnerGroup.Name == "Rotation About")
            {
              bpg = BuiltInParameterGroup.PG_ROTATION_ABOUT  ;
            }
            else if (def.OwnerGroup.Name == "Segments Fittings")
            {
              bpg = BuiltInParameterGroup.PG_SEGMENTS_FITTINGS  ;
            }
            else if (def.OwnerGroup.Name == "Slab Shape Edit")
            {
              bpg = BuiltInParameterGroup.PG_SLAB_SHAPE_EDIT  ;
            }
            else if (def.OwnerGroup.Name == "Split Profile Dimensions")
            {
              bpg = BuiltInParameterGroup.PG_SPLIT_PROFILE_DIMENSIONS  ;
            }
            else if (def.OwnerGroup.Name == "Stair Risers")
            {
              bpg = BuiltInParameterGroup.PG_STAIR_RISERS  ;
            }
            else if (def.OwnerGroup.Name == "Stair Stringers")
            {
              bpg = BuiltInParameterGroup.PG_STAIR_STRINGERS  ;
            }
            else if (def.OwnerGroup.Name == "Stair Treads")
            {
              bpg = BuiltInParameterGroup.PG_STAIR_TREADS  ;
            }
            else if (def.OwnerGroup.Name == "Stairs Calculator Rules")
            {
              bpg = BuiltInParameterGroup.PG_STAIRS_CALCULATOR_RULES  ;
            }
            else if (def.OwnerGroup.Name == "Stairs Open End Connection")
            {
              bpg = BuiltInParameterGroup.PG_STAIRS_OPEN_END_CONNECTION  ;
            }
            else if (def.OwnerGroup.Name == "Stairs Supports")
            {
              bpg = BuiltInParameterGroup.PG_STAIRS_SUPPORTS  ;
            }
            else if (def.OwnerGroup.Name == "Stairs Treads Risers")
            {
              bpg = BuiltInParameterGroup.PG_STAIRS_TREADS_RISERS  ;
            }
            else if (def.OwnerGroup.Name == "Stairs Winders")
            {
              bpg = BuiltInParameterGroup.PG_STAIRS_WINDERS  ;
            }
            else if (def.OwnerGroup.Name == "Structural")
            {
              bpg = BuiltInParameterGroup.PG_STRUCTURAL  ;
            }
            else if (def.OwnerGroup.Name == "Structural Analysis")
            {
              bpg = BuiltInParameterGroup.PG_STRUCTURAL_ANALYSIS  ;
            }
            else if (def.OwnerGroup.Name == "Support")
            {
              bpg = BuiltInParameterGroup.PG_SUPPORT  ;
            }
            else if (def.OwnerGroup.Name == "Systemtype Risedrop")
            {
              bpg = BuiltInParameterGroup.PG_SYSTEMTYPE_RISEDROP  ;
            }
            else if (def.OwnerGroup.Name == "Termination")
            {
              bpg = BuiltInParameterGroup.PG_TERMINTATION  ;
            }
            else if (def.OwnerGroup.Name == "Text")
            {
              bpg = BuiltInParameterGroup.PG_TEXT  ;
            }
            else if (def.OwnerGroup.Name == "Title")
            {
              bpg = BuiltInParameterGroup.PG_TITLE  ;
            }
            else if (def.OwnerGroup.Name == "Translation In")
            {
              bpg = BuiltInParameterGroup.PG_TRANSLATION_IN  ;
            }
            else if (def.OwnerGroup.Name == "Truss Family Bottom Chord")
            {
              bpg = BuiltInParameterGroup.PG_TRUSS_FAMILY_BOTTOM_CHORD  ;
            }
            else if (def.OwnerGroup.Name == "Truss Family Diag Web")
            {
              bpg = BuiltInParameterGroup.PG_TRUSS_FAMILY_DIAG_WEB  ;
            }
            else if (def.OwnerGroup.Name == "Truss Family Top Chord")
            {
              bpg = BuiltInParameterGroup.PG_TRUSS_FAMILY_TOP_CHORD  ;
            }
            else if (def.OwnerGroup.Name == "Truss Family Vert Web")
            {
              bpg = BuiltInParameterGroup.PG_TRUSS_FAMILY_VERT_WEB  ;
            }
            else if (def.OwnerGroup.Name == "View Camera")
            {
              bpg = BuiltInParameterGroup.PG_VIEW_CAMERA  ;
            }
            else if (def.OwnerGroup.Name == "View Extents")
            {
              bpg = BuiltInParameterGroup.PG_VIEW_EXTENTS  ;
            }
            else if (def.OwnerGroup.Name == "Visibility")
            {
              bpg = BuiltInParameterGroup.PG_VISIBILITY  ;
            }
            else
            {
              bpg = BuiltInParameterGroup.INVALID;
            }

            m_manager.AddParameter(def, bpg, false);

          }
          catch (System.Exception e)
          {
            MessageBox.Show(e.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            return false;
          }
        }
      }
      return true;
    }
  }
}
