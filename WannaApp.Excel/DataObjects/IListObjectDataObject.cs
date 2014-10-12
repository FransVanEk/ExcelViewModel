using System;
namespace WannaApp.Excel.DataObjects
{
    public interface IListObjectDataObject
    {
        object[,] AllValues { get; }
        object[,] DataValues { get;  }
        string[] HeaderValues { get;  }
  
    }
}
