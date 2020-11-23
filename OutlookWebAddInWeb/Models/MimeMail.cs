using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace OutlookWebAddInWeb.Models
{
  public class MimeMail
  {
    public string MessageID { get; set; }
    public bool IsValid()
    {
      return !string.IsNullOrEmpty(MessageID);
    }
  }
}