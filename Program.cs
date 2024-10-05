using Google.Apis.Auth.OAuth2;
using Google.Apis.Drive.v3;
using Google.Apis.Services;
using Google.Apis.Util.Store;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using YoutubeDLSharp;

class Program
{

    private static readonly string FFMpegPath = Path.Combine(Environment.CurrentDirectory, "bin");
    private static readonly string YTDLPath = Path.Combine(Environment.CurrentDirectory, "bin\\yt-dlp.exe");

    static string[] Scopes = { DriveService.Scope.DriveReadonly };
    static string ApplicationName = "Google Drive API .NET";
    static async Task Main(string[] args)
    {

        var ytdl = new YoutubeDL
        {
            YoutubeDLPath = YTDLPath,
            FFmpegPath = FFMpegPath,
            //OutputFolder = downloadedAudiosPath,
            //OutputFileTemplate = "%(title)s.%(ext)s"
        };
        string[] playlistUrls =
        {
            "https://www.youtube.com/playlist?list=PL4JCp4qfxq-oHhJDlj8vspSazWV7nWQX6",
            "https://www.youtube.com/playlist?list=PL4JCp4qfxq-pW0MncctTYbq0LO02TKLQn",
            "https://www.youtube.com/playlist?list=PL4JCp4qfxq-ryds1VN1NvODZFGbDHa3Hg",
            "https://www.youtube.com/playlist?list=PL4JCp4qfxq-q-rTihANSybJeipB6ERqiw",
            "https://www.youtube.com/playlist?list=PL4JCp4qfxq-rkyD27V9yod-uTSY1j3tGk",
            "https://www.youtube.com/playlist?list=PL4JCp4qfxq-pGzGfn1CZ1MI9SYkNGwDgj"
        };

        var excelData = new List<ExcelDto>();
        foreach (var playlistUrl in playlistUrls)
        {
            var playlistData = (await ytdl.RunVideoDataFetch(playlistUrl))?.Data;

            Console.WriteLine("Fetching playlist: " + playlistData.Title);

            // remove private video by filter title, if title is not null is public, else private
            var selectedVideos1 = playlistData.Entries.Select(x => new ExcelDto
            {
                Id = x.ID,
                Title = x.Title,
                YoutubeLink = $"https://www.youtube.com/watch?v={x.ID}"
            })
            .ToList();

            var selectedVideos = playlistData.Entries.Select(x => new ExcelDto
            {
                Id = x.ID,
                Title = x.Title,
                YoutubeLink = $"https://www.youtube.com/watch?v={x.ID}"
            })
            .Where(x => x.Title != "[Private video]")
            .ToList();

            if (selectedVideos1.Count != selectedVideos.Count)
            {
                Console.WriteLine( playlistData.Title + "có private video");
            }


            excelData.AddRange(selectedVideos);
        }

        var dic = GetDriveLink();
        foreach (var item in excelData)
        {
            if (dic.ContainsKey(item.Id))
            {
                item.DriveLink = dic[item.Id];
            }
        }


        string filePath = @"D:\Testing Youtube Subtitle - FFmpeg\Sư Hạnh Tuệ\output\Vấn đáp Phật Pháp.xlsx"; // Đường dẫn lưu file Excel
        ExportToExcel(excelData, filePath);
    }


    static void ExportToExcel(List<ExcelDto> excelData, string filePath)
    {
        // Tạo workbook mới
        IWorkbook workbook = new XSSFWorkbook();

        // Tạo sheet mới
        ISheet sheet = workbook.CreateSheet("Playlists");

        // Tạo hàng đầu tiên (header)
        IRow headerRow = sheet.CreateRow(0);
        headerRow.CreateCell(0).SetCellValue("Tiêu đề video");
        headerRow.CreateCell(1).SetCellValue("Link Youtube");
        headerRow.CreateCell(2).SetCellValue("Link Google Drive");

        // Tạo style cho hyperlink
        ICellStyle linkStyle = CreateHyperlinkStyle(workbook);

        // Duyệt qua dữ liệu và tạo các hàng tiếp theo
        for (int i = 0; i < excelData.Count; i++)
        {
            var data = excelData[i];
            IRow row = sheet.CreateRow(i + 1);  // hàng tiếp theo (bắt đầu từ 1 vì 0 là header)

            // Set tiêu đề video
            row.CreateCell(0).SetCellValue(data.Title);

            // Tạo cell cho link Youtube
            ICell youtubeCell = row.CreateCell(1);
            youtubeCell.SetCellValue(data.YoutubeLink); // Văn bản hiển thị cho link

            // Tạo hyperlink cho Youtube
            IHyperlink youtubeLink = workbook.GetCreationHelper().CreateHyperlink(HyperlinkType.Url);
            youtubeLink.Address = data.YoutubeLink; // Gán URL Youtube
            youtubeCell.Hyperlink = youtubeLink;
            youtubeCell.CellStyle = linkStyle; // Áp dụng style cho hyperlink

            // Tạo cell cho link Google Drive
            ICell driveCell = row.CreateCell(2);
            driveCell.SetCellValue(data.DriveLink); // Văn bản hiển thị cho link

            // Tạo hyperlink cho Google Drive
            IHyperlink driveLink = workbook.GetCreationHelper().CreateHyperlink(HyperlinkType.Url);
            driveLink.Address = data.DriveLink; // Gán URL Drive
            driveCell.Hyperlink = driveLink;
            driveCell.CellStyle = linkStyle; // Áp dụng style cho hyperlink
        }

        // Ghi workbook ra file
        using (FileStream stream = new FileStream(filePath, FileMode.Create, FileAccess.Write))
        {
            workbook.Write(stream);
        }
    }

    // Hàm tạo style cho hyperlink (chữ xanh, gạch chân)
    private static ICellStyle CreateHyperlinkStyle(IWorkbook workbook)
    {
        IFont hlinkFont = workbook.CreateFont();
        hlinkFont.Underline = FontUnderlineType.Single;
        hlinkFont.Color = IndexedColors.Blue.Index;

        ICellStyle linkStyle = workbook.CreateCellStyle();
        linkStyle.SetFont(hlinkFont);

        return linkStyle;
    }

    private static Dictionary<string, string> GetDriveLink()
    {
        UserCredential credential;

        using (var stream =
            new FileStream("D:\\Testing Youtube Subtitle - FFmpeg\\Sư Hạnh Tuệ\\credentials_youtube.json", FileMode.Open, FileAccess.Read))
        {
            string credPath = "D:\\Testing Youtube Subtitle - FFmpeg\\Sư Hạnh Tuệ\\token.json";
            credential = GoogleWebAuthorizationBroker.AuthorizeAsync(
                GoogleClientSecrets.Load(stream).Secrets,
                Scopes,
                "user",
                CancellationToken.None,
                new FileDataStore(credPath, true)).Result;
            Console.WriteLine("Credential file saved to: " + credPath);
        }

        // Tạo một đối tượng DriveService để gọi Google Drive API
        var service = new DriveService(new BaseClientService.Initializer()
        {
            HttpClientInitializer = credential,
            ApplicationName = ApplicationName,
        });

        // Folder ID của thư mục bạn muốn lấy thông tin file
        string parentFolderId = "1GGrk-q2hugxfyOpqxaV1CLphQ1707WGd";

        // Bước 1: Lấy tất cả các folder con trong folder cha
        var subfolders = ListFolders(service, parentFolderId);

        subfolders = subfolders.Where(x => x.Name.Contains("VẤN ĐÁP")).ToList();

        var dic = new Dictionary<string, string>();
        // Bước 2: Duyệt qua từng folder con và lấy tất cả các file trong từng folder
        foreach (var folder in subfolders)
        {
            Console.WriteLine($"Folder: {folder.Name}");
            var filesInFolder = ListFilesInFolder(service, folder.Id);
            foreach (var file in filesInFolder)
            {
                //Console.WriteLine($"  File: {ExtractIdentifier(file.Name)}");
                //Console.WriteLine($"  Link: {file.WebViewLink}");
                var videoId = ExtractIdentifier(file.Name);
                if (!dic.ContainsKey(videoId))
                {
                    dic.Add(videoId, file.WebViewLink);
                }
            }
        }

        return dic;
    }

    // Hàm lấy tất cả các folder con trong folder cha
    static IList<Google.Apis.Drive.v3.Data.File> ListFolders(DriveService service, string parentFolderId)
    {
        FilesResource.ListRequest listRequest = service.Files.List();
        listRequest.Q = $"'{parentFolderId}' in parents and mimeType = 'application/vnd.google-apps.folder' and trashed = false";
        listRequest.Fields = "nextPageToken, files(id, name)";

        var folders = listRequest.Execute().Files;
        return folders;
    }

    // Hàm lấy tất cả các file trong một folder con
    static IList<Google.Apis.Drive.v3.Data.File> ListFilesInFolder(DriveService service, string folderId)
    {
        List<Google.Apis.Drive.v3.Data.File> allFiles = new List<Google.Apis.Drive.v3.Data.File>();
        string pageToken = null;

        do
        {
            FilesResource.ListRequest listRequest = service.Files.List();
            listRequest.Q = $"'{folderId}' in parents and mimeType != 'application/vnd.google-apps.folder' and trashed = false";
            listRequest.Fields = "nextPageToken, files(id, name, webViewLink)";
            listRequest.PageToken = pageToken;

            var fileList = listRequest.Execute();

            if (fileList.Files != null && fileList.Files.Count > 0)
            {
                allFiles.AddRange(fileList.Files);
            }

            // Gán pageToken để lấy trang tiếp theo
            pageToken = fileList.NextPageToken;

        } while (pageToken != null);

        return allFiles;
    }

    public static string ExtractIdentifier(string input)
    {
        // Check if the input string is null or empty
        if (string.IsNullOrEmpty(input))
            return string.Empty;

        // Find the opening and closing brackets
        int startIndex = input.IndexOf('[');
        int endIndex = input.IndexOf(']', startIndex);

        // If brackets are found, extract the substring between them
        if (startIndex != -1 && endIndex != -1 && endIndex > startIndex)
        {
            string identifier = input.Substring(startIndex + 1, endIndex - startIndex - 1);

            // Check if the identifier is exactly 11 characters long
            if (identifier.Length == 11)
            {
                return identifier;
            }
        }

        // Return an empty string if no valid identifier is found
        return string.Empty;
    }
}


class ExcelDto
{
    public string Id { get; set; }
    public string Title { get; set; }
    public string YoutubeLink { get; set; }
    public string DriveLink { get; set; }
}