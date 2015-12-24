using System;
using System.Collections.Generic;
using System.IO;
using System.Xml.Serialization;

namespace TFSTasksInOutlook.Dataset
  {
  public class TfsTasksStorage
    {
    public string TfsUri { get; set; }
    public List<string> TfsProjects { get; set; }
    public List<WorkItemInfo> FavoriteWorkItems { get; set; }

    public TfsTasksStorage()
      {
      TfsUri = "";
      TfsProjects = new List<string>();
      FavoriteWorkItems = new List<WorkItemInfo>();
      }

    public static TfsTasksStorage Load()
      {
      var path = _GetStoragePath();
      if (!File.Exists(path)) return new TfsTasksStorage();
      var serializer = new XmlSerializer(typeof(TfsTasksStorage));
      using (TextReader reader = new StreamReader(path))
        {
        var res = serializer.Deserialize(reader) as TfsTasksStorage;
        if (res != null) return res;
        }
      return new TfsTasksStorage();
      }

    public void Save()
      {
      var path = _GetStoragePath();
      var serializer = new XmlSerializer(typeof(TfsTasksStorage));
      using (TextWriter writer = new StreamWriter(path))
        {
        serializer.Serialize(writer, this);
        }
      }

    private static string _GetStoragePath()
      {
      var path = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.CommonApplicationData), @"TfsTasksInOutlook\data.xml");
      var containingDirectory = Path.GetDirectoryName(path);
      if (containingDirectory != null && !Directory.Exists(containingDirectory))
        Directory.CreateDirectory(containingDirectory);
      return path;
      }
    }
  }
