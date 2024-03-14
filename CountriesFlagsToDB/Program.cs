using System;
using System.Data.SqlClient;
using System.IO;
using System.Net;

namespace CountriesFlagsToDB
{
    class Program
    {
        static void Main(string[] args)
        {
            // Database connection string
            string connectionString = "Data Source=.;Initial Catalog=niksaj;Integrated Security=True";

            // Array of country codes
            string[] countryCodes = "af,au,be,bn,cf,ci,do,sz,de,ht,ie,ki,lb,mw,mx,mm,ng,py,rw,sn,so,ch,tt,gb,ye,am,by,br,ca,cr,dm,ee,ge,gy,iq,ke,lv,mg,mu,mz,ne,pg,ru,sa,sb,se,to,ae,vn,ao,bh,bo,cv,co,cz,sv,fr,gt,in,jp,kw,li,mt,mn,nl,pk,pt,ws,sg,lk,tz,tv,vu,ar,bb,bw,cm,cg,dj,er,gm,gw,ir,kz,la,lu,mr,ma,ni,pa,ro,st,si,sr,tg,ua,ve,ad,bs,bt,bi,cn,cy,eg,fi,gd,is,jm,xk,ly,ml,mc,np,om,pl,vc,sl,es,tj,tm,uz,ag,bd,ba,kh,km,dk,gq,ga,gn,id,jo,kg,lt,mh,me,nz,pw,qa,sm,sk,sd,th,ug,va,al,at,bz,bg,td,hr,tl,et,gh,hn,il,kp,ls,my,fm,na,mk,pe,kn,rs,za,sy,tn,us,zm,dz,az,bj,bf,cl,cu,ec,fj,gr,hu,it,kr,lr,mv,md,nr,no,ph,lc,sc,ss,tw,tr,uy,zw".Split(',');

            foreach (string countryCode in countryCodes)
            {
                // URL of the flag image
                string imageUrl = "https://flagcdn.com/w320/" + countryCode.ToLower() + ".png";

                // Download the image
                byte[] imageData = DownloadImage(imageUrl);

                if (imageData != null)
                {
                    // Update database with the downloaded image
                    UpdateCountryFlagInDatabase(connectionString, countryCode.ToUpper(), imageData);
                    Console.WriteLine($"Flag for country code {countryCode.ToLower()} updated successfully.");
                }
                else
                {
                    Console.WriteLine($"Failed to download flag for country code {countryCode.ToUpper()}.");
                }
            }

            Console.WriteLine("All flag updates completed.");
        }

        static byte[] DownloadImage(string imageUrl)
        {
            try
            {
                using (WebClient webClient = new WebClient())
                {
                    byte[] imageData = webClient.DownloadData(imageUrl);
                    return imageData;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error downloading image: " + ex.Message);
                return null;
            }
        }

        static void UpdateCountryFlagInDatabase(string connectionString, string countryCode, byte[] imageData)
        {
            try
            {
                using (SqlConnection connection = new SqlConnection(connectionString))
                {
                    connection.Open();
                    string query = "UPDATE [dbo].[Country] SET [CountryFlagImage] = @ImageData WHERE [ISO_3166_code] = @CountryCode";
                    SqlCommand command = new SqlCommand(query, connection);
                    command.Parameters.AddWithValue("@ImageData", imageData);
                    command.Parameters.AddWithValue("@CountryCode", countryCode);
                    command.ExecuteNonQuery();
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error updating database: " + ex.Message);
            }
        }
    }
}
