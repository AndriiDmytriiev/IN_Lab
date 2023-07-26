// Decompiled with JetBrains decompiler
// Type: BI_CPV_tool.Program
// Assembly: BI-CPV-tool, Version=1.0.0.0, Culture=neutral, PublicKeyToken=null
// MVID: 15F98AF2-C907-4F3F-9C6A-820E171413A8
// Assembly location: C:\Users\49175\Documents\Version17\Version17\project.exe

using System;
using System.Windows.Forms;

namespace BI_CPV_tool
{
  internal static class Program
  {
    [STAThread]
    private static void Main()
    {
      Application.EnableVisualStyles();
      Application.SetCompatibleTextRenderingDefault(false);
      Application.Run((Form) new Form1());
    }
  }
}
