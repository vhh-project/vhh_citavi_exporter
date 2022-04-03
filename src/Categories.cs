using System;
using SwissAcademic.Citavi;
using SwissAcademic.Citavi.Shell;

public class Categories
{
	public static void Export(String path)
	{
		Project activeProject = Program.ActiveProjectShell.Project;
		string FOLDERPATH = path+ "\\categories.json";
		try
		{


			System.IO.File.WriteAllBytes(FOLDERPATH, new byte[0]);
			char[] charsToTrim = { ' ', '“', '”', };

			int count = 0;
			using (System.IO.StreamWriter file = new System.IO.StreamWriter(FOLDERPATH, true))
			{
				file.WriteLine("{\r\n\"categories\":[");

				foreach (Category cat in activeProject.AllCategories)
				{

					string parent;
					string parentClassifier;
					if (cat.IsRootCategory)
					{
						parent = "";
						parentClassifier = "";
					}
					else
					{
						parent = cat.ParentCategory.Name.ToString();
						parentClassifier = cat.Parent.Classification.ToString();
					}

					string temp = "";
					if (count > 0)
					{
						temp = ",{\r\n" +
                        "\"Index\":" + cat.Index + ",\r\n" +
						"\"Name\":\"" + JavaScriptStringEncode(cat.Name.Trim(charsToTrim),false) + "\",\r\n" +
						"\"Classifier\":\"" + cat.Classification + "\",\r\n" +
						"\"ParentName\":\"" + JavaScriptStringEncode(parent.Trim(charsToTrim), false) + "\",\r\n" +
						"\"ParentClassifier\":\"" + parentClassifier + "\",\r\n" +
						"\"Level\":" + cat.Level + "\r\n" +
						"}";
					}
					else
					{
						temp = "{\r\n" +
						"\"Index\":" + cat.Index + ",\r\n" +
						"\"Name\":\"" + JavaScriptStringEncode(cat.Name.Trim(charsToTrim), false) + "\",\r\n" +
						"\"Classifier\":\"" + cat.Classification + "\",\r\n" +
						"\"ParentName\":\"" + JavaScriptStringEncode(parent.Trim(charsToTrim),false) + "\",\r\n" +
						"\"ParentClassifier\":\"" + parentClassifier + "\",\r\n" +
						"\"Level\":" + cat.Level + "\r\n" +
						"}";
					}




					file.WriteLine(temp);

					count++;

				}

				file.WriteLine("]\r\n}");

				DebugMacro.WriteLine("Success");
			}
		}
		catch (Exception ex)
		{
			DebugMacro.WriteLine("Exception:" + ex.Message);
			DebugMacro.WriteLine(ex.StackTrace);
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
