using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

using System.Windows.Forms;
using HtmlAgilityPack;

namespace OLX_parser
{
    class Parse
    {
        string searchQuery;
        string region;
        string rubric; 
        string subrubric;
        List<string> listOfLinksOnOffers;
              
        DataGridView dataGrid;

        public Parse(string searchQuery, string region, string rubric, string subrubric, DataGridView newGrid)
        {
            this.searchQuery = searchQuery;
            this.region = region;
            this.rubric = rubric;
            this.subrubric = subrubric;           
            dataGrid = newGrid;

            listOfLinksOnOffers = new List<string>();


        }
        
        public async void parseAsync(CancellationToken token)
        {
            string url = Filters.gelLink(searchQuery, region, rubric, subrubric);
            string error = null;
            HtmlWeb web = new HtmlWeb();
            HtmlAgilityPack.HtmlDocument doc;

            doc = web.Load(url);

            try
            {
                error = doc.DocumentNode.SelectSingleNode("//*[@class='emptynew  large lheight18']/p/span").InnerText;

                if ((error == "Проверьте правильность написания или введите другие параметры поиска") || (error == "Перевірте правильність написання або введіть інші параметри пошуку"))
                {
                    MessageBox.Show("По данному запросу результатов нет!");
                }
            }
            catch(NullReferenceException)
            {
                do
                {
                    if (token.IsCancellationRequested)
                    {
                        return;
                    }
                    try
                    {
                        await Task.Run(() => parsePage(url, token));

                        url = doc.DocumentNode.SelectSingleNode("//*[@class='fbold next abs large']/a").Attributes[1].Value;
                        doc = web.Load(url);
                    }
                    catch (NullReferenceException)
                    {
                        return;
                    }
                } while (true);
            }  
        }


        private void parsePage(string link, CancellationToken token)
        {
            HtmlWeb web = new HtmlWeb(); 
            HtmlAgilityPack.HtmlDocument doc = web.Load(link);

            listOfLinksOnOffers.Clear();

            var listOfOffers = doc.DocumentNode.SelectNodes("//*[@class='wrap']/td");
            if (listOfOffers == null)
                return;

            foreach (var links in listOfOffers)
            {
                if (token.IsCancellationRequested)
                {
                    return;
                }
                var linkOnOffer = links.SelectSingleNode(".//td/a").Attributes[1].Value;
                listOfLinksOnOffers.Add(linkOnOffer);
            }

            foreach (var offer in listOfLinksOnOffers)
            {
                if (token.IsCancellationRequested)
                {
                    return;
                }

                doc = web.Load(offer);

                var name = doc.DocumentNode.SelectSingleNode("//*[@class = 'css-r9zjja-Text eu5v0x0']").InnerText;

                string price;

                try
                {
                    price = doc.DocumentNode.SelectSingleNode("//*[@class = 'css-okktvh-Text eu5v0x0']").InnerText;
                }
                catch (NullReferenceException Ex)
                {
                    price = "No price";
                }
                
                var published = doc.DocumentNode.SelectSingleNode("//*[@class = 'css-19yf5ek']").InnerText;

                var id = doc.DocumentNode.SelectSingleNode("//*[@class = 'css-9xy3gn-Text eu5v0x0']").InnerText;

                var description = doc.DocumentNode.SelectSingleNode("//*[@class = 'css-g5mtbi-Text']").InnerText;

                string pictureUrl;
                try
                {
                     pictureUrl = doc.DocumentNode.SelectSingleNode("//div[@class = 'swiper-zoom-container']/img").Attributes[0].Value;
                }
                catch(NullReferenceException Ex)
                {
                     pictureUrl = "No picture";
                }
                
                

                dataGrid.Invoke((MethodInvoker)delegate
                {
                    dataGrid.Rows.Add(name, description, offer, price, id, published, pictureUrl);
                });
            }

        }


    }
}
