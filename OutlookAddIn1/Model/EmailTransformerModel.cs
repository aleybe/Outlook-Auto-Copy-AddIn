using ConsoleApp4.Extensions;
using HtmlAgilityPack;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace ConsoleApp4.Model
{
    public class EmailTransformerModel : IEmailTransformerModel
    {
        public string ConvertTo(HtmlDocument data)
        {
            /// GET THE NODES WHICH MAKE UP THE EXPECTED RESULT
            var parentEmailBodyNodes = data.DocumentNode
                    .SelectSingleNode("//body") //This is reliable
                    .ChildNodes //get all child nodes
                    .FirstOrDefault(x => x.Name == "div")
                    .ChildNodes.TakeWhile(cur => cur.Name != "div");

            /// JOIN THOSE NODES INTO A PLAIN-TEXT VERSION
            //var body = string.Join("\n", parentEmailBodyNodes
            //                                ////.Select(x => WebUtility.HtmlDecode(x.InnerHtml.Replace("<br>","\n")))  
            //                                ////.Select(x => Regex.Replace(x,"<.+?>",""))
            //                                ////); 
            //                                .Select(x => WebUtility.HtmlDecode(x.InnerText))
            //                                );
            //                                //.Where(e => !string.IsNullOrWhiteSpace(e.InnerText))
            //                                //.Select(e => e.InnerText
            //                                //.Replace("&nbsp;", "")
            //                                //));

            //return body;
            /// JOIN THOSE NODES INTO A PLAIN-TEXT VERSION
            var body = string.Join("\n", parentEmailBodyNodes
                                            .Select(x =>
                                            {
                                                //Modify the innerHTML to replace specific tags
                                                x.InnerHtml = FormatInnerHtml(x);

                                                //The innertext will change based on the innerHtml Modification :)
                                                return x.InnerText;

                                            })

                                            );

            return body;
        }

        private string FormatInnerHtml(HtmlNode x)
        {

            string finalHtml = x.InnerHtml;

            finalHtml = finalHtml
                .Replace("<br>", "\n")
                //Add more .Replace( ) here


                ; // leave this semicolon alone :(

            return WebUtility.HtmlDecode(finalHtml);
        }
        //This isn't used so keep the exception there to break the program
        public HtmlDocument ConvertFrom(string data)
        {
            throw new NotImplementedException();
        }
    }

    public interface IEmailTransformerModel : IConvert<HtmlDocument, string>
    {
    }
}
