using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Net.Http.Formatting;
using System.Threading.Tasks;
using System.Web.Http.Description;
using System.Xml.Linq;
using Microsoft.Azure.WebJobs;
using Microsoft.Azure.WebJobs.Extensions.Http;
using Microsoft.Azure.WebJobs.Host;
using Microsoft.SharePoint.Client;
using OfficeDevPnP.Core.Utilities;
using Cumulus.Monads.Helpers;

namespace Cumulus.Monads.SharePoint
{
    public static class SetSiteColor
    {
        [FunctionName("SetSiteColor")]
        [ResponseType(typeof(SetSiteColorResponse))]
        [Display(Name = "Set color theme for the site", Description = "Define a color scheme for the site based off a base color.")]
        public static async Task<HttpResponseMessage> Run([HttpTrigger(AuthorizationLevel.Function, "post")]SetSiteColorRequest request, TraceWriter log)
        {
            if (!request.RGB.StartsWith("#")) request.RGB = "#" + request.RGB;
            if (request.RGB.Length != 7 && request.RGB.Length != 4)
            {
                return await Task.FromResult(new HttpResponseMessage(HttpStatusCode.BadGateway)
                {
                    Content = new ObjectContent<string>("Color was not in correct format. Use #ccc or #cccccc.", new JsonMediaTypeFormatter())
                });
            }

            try
            {
                string siteUrl = request.SiteURL;
                XDocument colorScheme = GenerateSPColor(request.RGB);
                
                MemoryStream spcolorStream = new MemoryStream();
                colorScheme.Save(spcolorStream);
                spcolorStream.Position = 0;

                var clientContext = await ConnectADAL.GetClientContext(siteUrl, log);
                var web = clientContext.Web;
                web.Lists.EnsureSiteAssetsLibrary();                
                const string fileName = "theme.spcolor";
                string relativeSiteUrl = UrlUtility.MakeRelativeUrl(siteUrl);
                string siteAssetsUrl = UrlUtility.Combine(relativeSiteUrl, "SiteAssets");
                var folder = web.GetFolderByServerRelativeUrl(siteAssetsUrl);

                var file = folder.UploadFile(fileName, spcolorStream, true);
                var fileUrl = UrlUtility.Combine(siteAssetsUrl, fileName);

                web.ApplyTheme(fileUrl, null, null, true);
                web.Update();
                clientContext.ExecuteQueryRetry();

                try
                {
                    file.DeleteObject();
                    clientContext.ExecuteQueryRetry();
                }
                catch (Exception)
                {
                    //Don't worry if deleting file fails
                }

                return await Task.FromResult(new HttpResponseMessage(HttpStatusCode.OK)
                {
                    Content = new ObjectContent<SetSiteColorResponse>(new SetSiteColorResponse { ColorSet = true }, new JsonMediaTypeFormatter())
                });
            }
            catch (Exception e)
            {
                log.Error($"Error:  {e.Message }\n\n{e.StackTrace}");
                return await Task.FromResult(new HttpResponseMessage(HttpStatusCode.ServiceUnavailable)
                {
                    Content = new ObjectContent<string>(e.Message, new JsonMediaTypeFormatter())
                });
            }
        }

        public class SetSiteColorRequest
        {
            [Required]
            [Display(Description = "URL of site")]
            public string SiteURL { get; set; }

            [Required]
            [Display(Description = "RGB color in hex notation. Eg. #00ff00")]
            public string RGB { get; set; }
        }

        public class SetSiteColorResponse
        {
            [Display(Description = "True if color was set")]
            public bool ColorSet { get; set; }
        }

        private static XDocument GenerateSPColor(string newHexColor)
        {
            XDocument doc = XDocument.Parse(_baseXml);
            var allSlots = doc.Descendants().Where(d => d.Name.LocalName == "color").ToList();

            var baseColorNode = allSlots.FirstOrDefault(n => n.FirstAttribute.Value == "AccentText");
            string hexColor = baseColorNode.Attributes().Where(a => a.Name == "value").Select(a => a.Value).Single();

            Color baseColor = ColorTranslator.FromHtml("#" + hexColor);
            Color newColor = ColorTranslator.FromHtml(newHexColor);

            List<string> consistentSlots = new List<string> { "ErrorText", "SearchURL", "TileText", "TileBackgroundOverlay" };

            HSLColor hslBaseColor = HSLColor.FromRgbColor(baseColor);
            HSLColor hslNewColor = HSLColor.FromRgbColor(newColor);
            foreach (XElement slotColor in allSlots)
            {
                var slotHexColor = slotColor.Attributes().Where(a => a.Name == "value").Select(a => a.Value).Single();
                Color currentColor = ColorTranslator.FromHtml("#" + slotHexColor);
                HSLColor hslCurrentColor = HSLColor.FromRgbColor(currentColor);
                string slotName = slotColor.Attributes().Where(a => a.Name == "name").Select(a => a.Value).Single();

                if (!consistentSlots.Contains(slotName) && hslCurrentColor.Hue != 0.0 && hslCurrentColor.Saturation != 0.0)
                {
                    if (ColorsEqualIgnoringOpacity(baseColor, currentColor))
                    {
                        slotColor.Attributes().Single(a => a.Name == "value").Value = ColorTranslator.ToHtml(newColor).Trim('#');
                    }
                    else
                    {
                        double hueDifference = (double)hslCurrentColor.Hue - (double)hslBaseColor.Hue;
                        float hue = hslNewColor.Hue + (float)hueDifference;
                        float luminance = hslNewColor.Luminance;
                        float saturation = hslNewColor.Luminance + (hslCurrentColor.Luminance - hslBaseColor.Luminance);
                        HSLColor hslCalcColor = new HSLColor(hue, saturation, luminance);

                        Color calcColor = Color.FromArgb((int)currentColor.A, hslCalcColor.ToRgbColor());
                        slotColor.Attributes().Single(a => a.Name == "value").Value = ColorTranslator.ToHtml(calcColor).Trim('#');
                    }
                }
            }

            return doc;
        }

        private static bool ColorsEqualIgnoringOpacity(Color a, Color b)
        {
            if ((int)a.R == (int)b.R && (int)a.G == (int)b.G)
                return (int)a.B == (int)b.B;
            return false;
        }


        private static readonly string _baseXml = @"<?xml version=""1.0"" encoding=""utf-8""?>
<s:colorPalette isInverted=""false"" previewSlot1=""BackgroundOverlay"" previewSlot2=""BodyText"" previewSlot3=""AccentText"" xmlns:s=""http://schemas.microsoft.com/sharepoint/"">
    <s:color name=""BodyText"" value=""444444"" />
    <s:color name=""SubtleBodyText"" value=""777777"" />
    <s:color name=""StrongBodyText"" value=""262626"" />
    <s:color name=""DisabledText"" value=""B1B1B1"" />
    <s:color name=""SiteTitle"" value=""262626"" />
    <s:color name=""WebPartHeading"" value=""444444"" />
    <s:color name=""ErrorText"" value=""BF0000"" />
    <s:color name=""AccentText"" value=""0072C6"" />
    <s:color name=""SearchURL"" value=""338200"" />
    <s:color name=""Hyperlink"" value=""0072C6"" />
    <s:color name=""Hyperlinkfollowed"" value=""663399"" />
    <s:color name=""HyperlinkActive"" value=""004D85"" />
    <s:color name=""CommandLinks"" value=""666666"" />
    <s:color name=""CommandLinksSecondary"" value=""262626"" />
    <s:color name=""CommandLinksHover"" value=""0072C6"" />
    <s:color name=""CommandLinksPressed"" value=""004D85"" />
    <s:color name=""CommandLinksDisabled"" value=""B1B1B1"" />
    <s:color name=""BackgroundOverlay"" value=""D8FFFFFF"" />
    <s:color name=""DisabledBackground"" value=""FDFDFD"" />
    <s:color name=""PageBackground"" value=""FFFFFF"" />
    <s:color name=""HeaderBackground"" value=""D8FFFFFF"" />
    <s:color name=""FooterBackground"" value=""D8FFFFFF"" />
    <s:color name=""SelectionBackground"" value=""7F9CCEF0"" />
    <s:color name=""HoverBackground"" value=""7FCDE6F7"" />
    <s:color name=""RowAccent"" value=""0072C6"" />
    <s:color name=""StrongLines"" value=""92C0E0"" />
    <s:color name=""Lines"" value=""ABABAB"" />
    <s:color name=""SubtleLines"" value=""C6C6C6"" />
    <s:color name=""DisabledLines"" value=""E1E1E1"" />
    <s:color name=""AccentLines"" value=""2A8DD4"" />
    <s:color name=""DialogBorder"" value=""F0F0F0"" />
    <s:color name=""Navigation"" value=""666666"" />
    <s:color name=""NavigationAccent"" value=""0072C6"" />
    <s:color name=""NavigationHover"" value=""0072C6"" />
    <s:color name=""NavigationPressed"" value=""004D85"" />
    <s:color name=""NavigationHoverBackground"" value=""7FCDE6F7"" />
    <s:color name=""NavigationSelectedBackground"" value=""C6EFEFEF"" />
    <s:color name=""EmphasisText"" value=""FFFFFF"" />
    <s:color name=""EmphasisBackground"" value=""0072C6"" />
    <s:color name=""EmphasisHoverBackground"" value=""0067B0"" />
    <s:color name=""EmphasisBorder"" value=""0067B0"" />
    <s:color name=""EmphasisHoverBorder"" value=""004D85"" />
    <s:color name=""SubtleEmphasisText"" value=""666666"" />
    <s:color name=""SubtleEmphasisCommandLinks"" value=""262626"" />
    <s:color name=""SubtleEmphasisBackground"" value=""F1F1F1"" />
    <s:color name=""TopBarText"" value=""666666"" />
    <s:color name=""TopBarBackground"" value=""C6EFEFEF"" />
    <s:color name=""TopBarHoverText"" value=""333333"" />
    <s:color name=""TopBarPressedText"" value=""004D85"" />
    <s:color name=""HeaderText"" value=""444444"" />
    <s:color name=""HeaderSubtleText"" value=""777777"" />
    <s:color name=""HeaderDisableText"" value=""B1B1B1"" />
    <s:color name=""HeaderNavigationText"" value=""666666"" />
    <s:color name=""HeaderNavigationHoverText"" value=""0072C6"" />
    <s:color name=""HeaderNavigationPressedText"" value=""004D85"" />
    <s:color name=""HeaderNavigationSelectedText"" value=""0072C6"" />
    <s:color name=""HeaderLines"" value=""ABABAB"" />
    <s:color name=""HeaderStrongLines"" value=""92C0E0"" />
    <s:color name=""HeaderAccentLines"" value=""2A8DD4"" />
    <s:color name=""HeaderSubtleLines"" value=""C6C6C6"" />
    <s:color name=""HeaderDisabledLines"" value=""E1E1E1"" />
    <s:color name=""HeaderDisabledBackground"" value=""FDFDFD"" />
    <s:color name=""HeaderFlyoutBorder"" value=""D1D1D1"" />
    <s:color name=""HeaderSiteTitle"" value=""262626"" />
    <s:color name=""SuiteBarBackground"" value=""0072C6"" />
    <s:color name=""SuiteBarHoverBackground"" value=""4B9BD7"" />
    <s:color name=""SuiteBarText"" value=""FFFFFF"" />
    <s:color name=""SuiteBarDisabledText"" value=""92C0E0"" />
    <s:color name=""ButtonText"" value=""444444"" />
    <s:color name=""ButtonDisabledText"" value=""B1B1B1"" />
    <s:color name=""ButtonBackground"" value=""FDFDFD"" />
    <s:color name=""ButtonHoverBackground"" value=""E6F2FA"" />
    <s:color name=""ButtonPressedBackground"" value=""92C0E0"" />
    <s:color name=""ButtonDisabledBackground"" value=""FDFDFD"" />
    <s:color name=""ButtonBorder"" value=""ABABAB"" />
    <s:color name=""ButtonHoverBorder"" value=""92C0E0"" />
    <s:color name=""ButtonPressedBorder"" value=""2A8DD4"" />
    <s:color name=""ButtonDisabledBorder"" value=""E1E1E1"" />
    <s:color name=""ButtonGlyph"" value=""666666"" />
    <s:color name=""ButtonGlyphActive"" value=""444444"" />
    <s:color name=""ButtonGlyphDisabled"" value=""C6C6C6"" />
    <s:color name=""TileText"" value=""FFFFFF"" />
    <s:color name=""TileBackgroundOverlay"" value=""7F000000"" />
    <s:color name=""ContentAccent1"" value=""0072C6"" />
    <s:color name=""ContentAccent2"" value=""00485B"" />
    <s:color name=""ContentAccent3"" value=""288054"" />
    <s:color name=""ContentAccent4"" value=""767956"" />
    <s:color name=""ContentAccent5"" value=""ED0033"" />
    <s:color name=""ContentAccent6"" value=""682A7A"" />
</s:colorPalette>
";
    }
}
