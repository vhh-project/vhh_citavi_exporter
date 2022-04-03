using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using SwissAcademic.Citavi;
using SwissAcademic.Citavi.Shell;
using System.Net;
using System.IO;

public class Papers
{
	public static void Export(String path)
	{
		Project activeProject = Program.ActiveProjectShell.Project;
		string FOLDERPATH = path+ "\\references.json";
        System.IO.Directory.CreateDirectory(path + "\\Covers");

        try
		{
			
			System.IO.File.WriteAllBytes(FOLDERPATH, new byte[0]);

			char[] charsToTrim = { ' ', '“', '”', };
			int count = 0;
			using (System.IO.StreamWriter file = new System.IO.StreamWriter(FOLDERPATH, true))
			{
				file.WriteLine("{\r\n\"references\":[");

			    var customFields =
			        activeProject.ProjectSettings.CustomFields.ToDictionary((setting) => (setting.Property.PropertyId.ToString()));
				
				foreach (Reference reference in activeProject.References)
				{

                    var cover = reference.CoverPath;
                    if (cover != null && cover.ToString() != "")
                    {
                        try
                        {
                            using (WebClient client = new WebClient())
                            {
                                //client.DownloadFile(loc.Address.ToString(), FOLDERPATH + saveFilename);

                                var data = client.DownloadData(cover.Resolve(true).ToString());
                                File.WriteAllBytes(path + "\\Covers\\" + cover.ToString(), data);
                            }
                        }
                        
                        catch (WebException ex)
                        {
                            DebugMacro.WriteLine("Exception while downloading " + reference.Id.ToString());
                        }
                    }

					string temp = "";
					string allProps = "";
				    
					PropertyInfo[] properties = reference.GetType().GetProperties();

					foreach (var prop in properties)
					{
					    var valueIsText = true;
                        string key = prop.Name;
					    string value = JavaScriptStringEncode(prop.GetValue(reference).ToStringSafe(),false);

                        if (prop.Name.StartsWith("CustomField"))
                        {
                            key = customFields[prop.Name].LabelText;
                        }
                        else if (prop.Name == "Groups")
                        {
                            valueIsText = false;
                            var v = (ReferenceGroupCollection)prop.GetValue(reference);
                            if (v.Count == 0) value = "[]";
                            else value = "[" + v.Select(c => "\"" + JavaScriptStringEncode(c.FullName, false) + "\"").Aggregate((i, j) => i + "," + j) + "]";
                        }
                        else if (prop.Name == "Keywords")
                        {
                            valueIsText = false;
                            var v = (ReferenceKeywordCollection)prop.GetValue(reference);
                            if (v.Count == 0) value = "[]";
                            else value = "[" + v.Select(c => "\"" + JavaScriptStringEncode(c.Name, false) + "\"").Aggregate((i, j) => i + "," + j) + "]";
                        }
                        else if (prop.Name == "Categories")
                        {
                            valueIsText = false;
                            var v = (ReferenceCategoryCollection)prop.GetValue(reference);
                            if (v.Count == 0) value = "[]";
                            else value = "[" + v.Select(c => "\"" + JavaScriptStringEncode(c.Classification, false) + "\"").Aggregate((i, j) => i + "," + j) + "]";
                        }
                        else if (prop.Name == "Abstract")
                        {
                            value = JavaScriptStringEncode(reference.Abstract.Text, false);
                        }
                        else if (prop.Name == "EntityLinks")
                        {
                            valueIsText = false;
                            var v = ((IEnumerable<EntityLink>)prop.GetValue(reference)).ToList();
                            if (v.Count == 0) value = "[]";
                            else value = "[" + v.Select(i => $"{{\"Id\":\"{i.Id.ToString()}\",\"Text\":\"{JavaScriptStringEncode(i.ToString(), false)}\"}}")
                                    .Aggregate((i, j) => i + "," + j) + "]";
                        }
                        else if (prop.Name == "Locations")
                        {
                            valueIsText = false;
                            var v = ((IEnumerable<Location>)prop.GetValue(reference)).ToList();
                            if (v.Count == 0) value = "[]";
                            else value = "[" + v.Select(i => $"{{\"Id\":\"{i.Id.ToString()}\",\"FullName\":\"{JavaScriptStringEncode(i.FullName, false)}\",\"Address\":\"{JavaScriptStringEncode(i.Address, false)}\"}}")
                                                .Aggregate((i, j) => i + "," + j) + "]";
                        }

                            //if (valueIsText) value = JavaScriptStringEncode(value.Trim(charsToTrim),false);

                            if (!prop.Name.Equals("FormattedReference") 
                            && !prop.Name.Equals("IsPropertyChangeNotificationSuspended") 
                            && !prop.Name.Equals("StaticIds")
                            && !prop.Name.Equals("TableOfContents"))
						{
							if (!allProps.Equals(""))
							{
								allProps += ",\r\n";
							}
                            
							if (valueIsText) allProps += "\""+key+"\":\"" +value+ "\"";
                            else allProps += "\"" + key + "\":" + value + "";
                        }

					}

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
		    System.Windows.Forms.MessageBox.Show("Exception:" + e.Message);
		    System.Windows.Forms.MessageBox.Show(e.StackTrace);
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
