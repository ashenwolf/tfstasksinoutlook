using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.IO;
using System.Linq;
using System.Text;
using System.Xml.Serialization;

namespace TFSTasksInOutlook.Dataset
  {
  public class TFSTasksStorage
    {
    public string TfsUri { get; set; }
    public List<string> TfsProjects { get; set; }
    public List<WorkItemInfo> FavoriteWorkItems { get; set; }

    public TFSTasksStorage()
      {
      TfsUri = "";
      TfsProjects = new List<string>();
      FavoriteWorkItems = new List<WorkItemInfo>();
      }

    public void Load()
      {
      var path = _GetStoragePath();
      if (File.Exists(path))
        {
        XmlSerializer serializer = new XmlSerializer(typeof(TFSTasksStorage));
        using (TextReader reader = new StreamReader(path))
          {
          var res = serializer.Deserialize(reader) as TFSTasksStorage;
          if (res != null)
            {
            this.TfsUri = res.TfsUri;
            this.TfsProjects = res.TfsProjects;
            this.FavoriteWorkItems = res.FavoriteWorkItems;
            }
          }
        }
      }

    public void Save()
      {
      var path = _GetStoragePath();
      XmlSerializer serializer = new XmlSerializer(typeof(TFSTasksStorage));
      using (TextWriter writer = new StreamWriter(path))
        {
        serializer.Serialize(writer, this);
        }
      }

    private string _GetStoragePath()
      {
      var path = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.CommonApplicationData), @"TfsTasksInOutlook\data.xml");
      if (!Directory.Exists(Path.GetDirectoryName(path)))
      Directory.CreateDirectory(Path.GetDirectoryName(path));
      return path;
      }
    }
  }
