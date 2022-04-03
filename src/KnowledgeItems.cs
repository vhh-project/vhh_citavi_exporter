using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using SwissAcademic.Citavi;
using SwissAcademic.Citavi.Shell;
using System.Net;
using System.IO;

public class KnowledgeItems
{
	public static void Export(String path)
	{
		Project activeProject = Program.ActiveProjectShell.Project;
		string FOLDERPATH = path+"\\knowledgeItems.json";
        System.IO.Directory.CreateDirectory(path + "\\KnowledgeItemFiles");

        try
        {
	
			System.IO.File.WriteAllBytes(FOLDERPATH, new byte[0]);

			int count = 0;
			char[] charsToTrim = {' ','“', '”',};
			using (System.IO.StreamWriter file = new System.IO.StreamWriter(FOLDERPATH, true))
			{
				file.WriteLine("{\r\n\"knowledgeItems\":[");

				foreach (KnowledgeItem item in activeProject.AllKnowledgeItemsFiltered)
				{
					if(item.KnowledgeItemType == KnowledgeItemType.File)
                    {
                        try
                        {
                            using (WebClient client = new WebClient())
                            {
                                //client.DownloadFile(loc.Address.ToString(), FOLDERPATH + saveFilename);

                                var data = client.DownloadData(item.Address.Resolve(true).ToString());
                                File.WriteAllBytes(path + "\\KnowledgeItemFiles\\" + item.Address.ToString(), data);
                            }
                        }

                        catch (WebException ex)
                        {
                            DebugMacro.WriteLine("Exception while downloading " + item.Address.ToString());
                        }

                    }
                    PropertyInfo[] properties = item.GetType().GetProperties();

					
					string allProps = "";

					foreach (var prop in properties)
					{
					    var valueIsText = true;
					    string key = prop.Name;
					    string value = JavaScriptStringEncode(prop.GetValue(item).ToStringSafe(),false);

					    if (prop.Name == "EntityLinks")
					    {
					        valueIsText = false;
					        var v = ((IEnumerable<EntityLink>) prop.GetValue(item)).ToList();
					        if (v.Count == 0) value = "[]";

					        else
					        {
					            value = "[";
					            foreach (var link in v)
					            {
					                if (link.Target is Annotation ann)
					                {
                                        if (value != "[") { value += ","; }

					                    value += $"{{\"Id\":\"{ann.Id.ToString()}\"";
					                    value += $",\"Location_Id\":\"{ann.Location.Id.ToString()}\"";
					                    value += $",\"Location_FullName\":\"{JavaScriptStringEncode(ann.Location.FullName,false)}\"";

					                    var q = ann.Quads.Select(i=> $"{{\"IsContainer\":\"{i.IsContainer}\",\"Page_Idx\":\"{i.PageIndex}\",\"X1\":\"{i.X1}\",\"X2\":\"{i.X2}\",\"Y1\":\"{i.Y1}\",\"Y2\":\"{i.Y2}\"}}").Aggregate((i, j) => i + "," + j);
					                    value += $",\"Quads\":[{q}]";
					                    value += "}";
					                }
					            }

					            value += "]";
					        }
					    }
					    else if (prop.Name == "Groups")
					    {
					        valueIsText = false;
					        var v = (KnowledgeItemGroupCollection)prop.GetValue(item);
					        if (v.Count == 0) value = "[]";
					        else value = "[" + v.Select(c => "\"" + JavaScriptStringEncode(c.FullName,false) + "\"").Aggregate((i, j) => i + "," + j) + "]";
					    }
					    else if (prop.Name == "Categories")
					    {
					        valueIsText = false;
					        var v = (KnowledgeItemCategoryCollection)prop.GetValue(item);
					        if (v.Count == 0) value = "[]";
					        else value = "[" + v.Select(c => "\"" + JavaScriptStringEncode(c.Classification,false) + "\"").Aggregate((i, j) => i + "," + j) + "]";
					    }
					    else if (prop.Name == "Keywords")
					    {
					        valueIsText = false;
					        var v = (KnowledgeItemKeywordCollection)prop.GetValue(item);
					        if (v.Count == 0) value = "[]";
					        else value = "[" + v.Select(c => "\"" + JavaScriptStringEncode(c.Name,false) + "\"").Aggregate((i, j) => i + "," + j) + "]";
					    }

                        if (!prop.Name.Equals("IsPropertyChangeNotificationSuspended")
                            && !prop.Name.Equals("StaticIds")
                            && !prop.Name.Equals("TextRtf"))
						{
							if (!allProps.Equals(""))
							{
								allProps += ",\r\n";
							}

						    //if (valueIsText) value = value.Trim(charsToTrim).Replace("\"", "\\\"");

						    if (valueIsText) allProps += "\"" + key + "\":\"" + value + "\"";
						    else allProps += "\"" + key + "\":" + value + "";
						}

                    }



					string temp = "";
					if (count > 0)
					{
						temp = ",{\r\n" +
						allProps + "\r\n"+
						"}";
					}
					else
					{
						temp = "{\r\n" +
						allProps + "\r\n" +
						"}";
					}



					file.WriteLine(temp);

					count++;
				}


				file.WriteLine("]\r\n}");
			}
			DebugMacro.WriteLine("Success");
		}
		catch (Exception e)
		{
			DebugMacro.WriteLine("Exception:" + e.Message);
			DebugMacro.WriteLine(e.StackTrace);
		}
	}

    public static string JavaScriptStringEncode(string value, bool addDoubleQuotes)
    {
        if (string.IsNullOrEmpty(value))
            return addDoubleQuotes ? "\"\"" : string.Empty;

        int len = value.Length;
        bool needEncode = false;
        char c;
        for (int i = 0; i < len; i++)
        {
            c = value[i];

            if (c >= 0 && c <= 31 || c == 34 || c == 39 || c == 60 || c == 62 || c == 92)
            {
                needEncode = true;
                break;
            }
        }

        if (!needEncode)
            return addDoubleQuotes ? "\"" + value + "\"" : value;

        var sb = new System.Text.StringBuilder();
        if (addDoubleQuotes)
            sb.Append('"');

        for (int i = 0; i < len; i++)
        {
            c = value[i];
            if (c >= 0 && c <= 7 || c == 11 || c >= 14 && c <= 31 || c == 39 || c == 60 || c == 62)
                sb.AppendFormat("\\u{0:x4}", (int)c);
            else switch ((int)c)
                {
                    case 8:
                        sb.Append("\\b");
                        break;

                    case 9:
                        sb.Append("\\t");
                        break;

                    case 10:
                        sb.Append("\\n");
                        break;

                    case 12:
                        sb.Append("\\f");
                        break;

                    case 13:
                        sb.Append("\\r");
                        break;

                    case 34:
                        sb.Append("\\\"");
                        break;

                    case 92:
                        sb.Append("\\\\");
                        break;

                    default:
                        sb.Append(c);
                        break;
                }
        }

        if (addDoubleQuotes)
            sb.Append('"');

        return sb.ToString();
    }
}
