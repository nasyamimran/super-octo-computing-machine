using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.IO;
using System.Data;
using System.Data.SqlClient;
using System.Text;
using Microsoft.Office.Interop.Word;
using Microsoft.Office;
using System.Configuration;
namespace DocProject
{
    public partial class WebForm1 : System.Web.UI.Page
    {
        string conStr = @"Data Source=DESKTOP-QITDMUU;Database=DocProj;Integrated Security=true;";


        protected void Page_Load(object sender, EventArgs e)
        {
           
        }

        


        protected void Button2_Click(object sender, EventArgs e)
        {
            int id=0;
            string[] words = new string[10];
            int[] freq = new int[10];



            string filePath = FileUpload1.PostedFile.FileName;
            string filename = Path.GetFileName(filePath);
            string ext = Path.GetExtension(filename);
            string contenttype = String.Empty;
            byte[] documentContent = null;
            if (ext == ".txt")
            {
                Stream fs = FileUpload1.PostedFile.InputStream;
                BinaryReader br = new BinaryReader(fs);
                documentContent = br.ReadBytes((Int32)fs.Length);
            }
            Application application = new Application();
            if (ext == ".doc" || ext == ".docx")
            {
              





                Document document = application.Documents.Open("C:\\Users\\pumpk\\Desktop\\" + filePath);
                int count = document.Words.Count;
                string text = "";
                for (int j = 1; j <= count; j++)
                {
                    // Write the word.
                    text = text + " " + document.Words[j].Text;

                }
                // Close word.

                 documentContent = System.Text.Encoding.UTF8.GetBytes(text);
            }
            IDictionary<string, int> rank = processing(documentContent);
            int i = 0;
            foreach(KeyValuePair<string, int> item in rank)
            {
                freq[i] = item.Value;
                words[i] = item.Key;
                i++;
            }


            using (SqlConnection cn = new SqlConnection(conStr))
            {
                SqlCommand cmd = new SqlCommand("SaveDocument", cn);

                String query = "INSERT INTO [DocProj].[dbo].[WordTable] (word,frequency,Doc_ID) VALUES (@word,@frequency,@Doc_ID)";

             



               // SqlCommand cmd2 = new SqlCommand("SAVEWORD", cn);
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.Parameters.Add("@Doc", SqlDbType.VarBinary).Value = documentContent;
                cn.Open();
                cmd.ExecuteNonQuery();
                using (SqlCommand command = new SqlCommand("SELECT ID FROM Doc_Table", cn))
                {
                    
                    using (SqlDataReader reader = command.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            string var = reader[0].ToString();
                            id = Int32.Parse(var);
                           
                        }
                    }
                }

                for (int j = 0; j < 10; j++)
                {
                    using (SqlCommand command2 = new SqlCommand(query, cn))
                    {




                        // Check Error


                        //command2.CommandType = CommandType.StoredProcedure;
                        command2.Parameters.Add("@word", SqlDbType.VarChar).Value = words[j];
                        command2.Parameters.Add("@frequency", SqlDbType.Int).Value = freq[j];
                        command2.Parameters.Add("@Doc_ID", SqlDbType.Int).Value = id;
                        command2.ExecuteNonQuery();



                    }
                }






            }
       
            application.Quit();

        }
        protected IDictionary<string, int> processing(byte[] documentContent)
        {
            IDictionary<string, int> Ranks = new Dictionary<string, int>();
            string inputString = Encoding.UTF8.GetString(documentContent, 0, documentContent.Length);
            inputString = inputString.ToLower();

            // Define characters to strip from the input and do it
            string[] stripChars = { ";", ",", ".", "-", "_", "^", "(", ")", "[", "]",
                        "0", "1", "2", "3", "4", "5", "6", "7", "8", "9", "\n", "\t", "\r" };
            foreach (string character in stripChars)
            {
                inputString = inputString.Replace(character, "");
            }

            // Split on spaces into a List of strings
            List<string> wordList = inputString.Split(' ').ToList();

            // Define and remove stopwords
            string[] stopwords = new string[] { "and", "the", "she", "for", "this", "you", "but", "are", "that", "how", "who" };
            foreach (string word in stopwords)
            {
                // While there's still an instance of a stopword in the wordList, remove it.
                // If we don't use a while loop on this each call to Remove simply removes a single
                // instance of the stopword from our wordList, and we can't call Replace on the
                // entire string (as opposed to the individual words in the string) as it's
                // too indiscriminate (i.e. removing 'and' will turn words like 'bandage' into 'bdage'!)
                while (wordList.Contains(word))
                {
                    wordList.Remove(word);
                }
            }

            // Create a new Dictionary object
            Dictionary<string, int> dictionary = new Dictionary<string, int>();

            // Loop over all over the words in our wordList...
            foreach (string word in wordList)
            {
                // If the length of the word is at least three letters...
                if (word.Length >= 3)
                {
                    // ...check if the dictionary already has the word.
                    if (dictionary.ContainsKey(word))
                    {
                        // If we already have the word in the dictionary, increment the count of how many times it appears
                        dictionary[word]++;
                    }
                    else
                    {
                        // Otherwise, if it's a new word then add it to the dictionary with an initial count of 1
                        dictionary[word] = 1;
                    }

                } // End of word length check

            } // End of loop over each word in our input

            // Create a dictionary sorted by value (i.e. how many times a word occurs)
            var sortedDict = (from entry in dictionary orderby entry.Value descending select entry).ToDictionary(pair => pair.Key, pair => pair.Value);

            // Loop through the sorted dictionary and output the top 10 most frequently occurring words
            int count = 1;
           
            foreach (KeyValuePair<string, int> pair in sortedDict)
            {
                // Output the most frequently occurring words and the associated word counts
                
                Ranks.Add(pair.Key, pair.Value);
                count++;

                // Only display the top 10 words then break out of the loop!
                if (count > 10)
                {
                    count = 1;
                    break;
                }
            }

            // Wait for the user to press a key before exiting

            return Ranks;
        }
        protected void GridView1_SelectedIndexChanged(object sender, EventArgs e)
        {
            /*int id = int.Parse((sender as LinkButton).CommandArgument);
            string embed = "<object data=\"{0}{1}\" type=\"application/vnd.ms-word\" width =\"500px\" height=\"600px\">";
         
            embed += "</object>";
            ltEmbed.Text = string.Format(embed, ResolveUrl("~/FileCS.ashx?Id="), id);*/

            GridViewRow row = (sender as LinkButton).NamingContainer as GridViewRow;
            divHtmlContent.Visible = true;
            divHtmlContent.InnerHtml = (row.FindControl("hfHtmlContent") as HiddenField).Value;
        }

        protected void LinkButton1_Click(object sender, EventArgs e)
        {

        }
    }
}