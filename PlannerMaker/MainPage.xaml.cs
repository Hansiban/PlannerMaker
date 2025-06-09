using OfficeOpenXml;
using PlannerMaker.Models;

namespace PlannerMaker;

public partial class MainPage : ContentPage
{
    #region Property

    //강의목표
    private const int MaxObjectives = 4;
    private readonly List<Entry> _lessonObjectiveEntries = [];

    //강의계획
    private const int MaxLessonPlans = 5;
    private readonly List<(DatePicker date, Entry plan, Entry note)> _lessonPlanEntries = [];

    public string AcademyName => AcademyNameEntry.Text;

    public DateTime ClassTime
    {
        get
        {
            int year = int.TryParse(YearPicker.SelectedItem?.ToString(), out var y) ? y : DateTime.Now.Year;
            int month = int.TryParse(MonthPicker.SelectedItem?.ToString(), out var m) ? m : DateTime.Now.Month;
            return new DateTime(year, month, 1);
        }
    }

    public string Subjects => SubjectsEntry.Text;
    public string Instructor => InstructorEntry.Text;

    public List<string> LessonObjectives => _lessonObjectiveEntries.Select(x => x.Text).ToList();

    public List<LessonPlan> LessonPlans => _lessonPlanEntries.Select(x => new LessonPlan
    {
        LessonDate = x.date.Date,
        PlanDetail = x.plan.Text?.Trim(),
        Note = x.note.Text?.Trim()
    }).ToList();

    #endregion Property

    public MainPage()
    {
        ExcelPackage.License.SetNonCommercialOrganization("PlannerMaker");
        InitializeComponent();
        InitYearMonthPickers();
    }

    /// <summary>
    /// Init Pickers
    /// </summary>
    private void InitYearMonthPickers()
    {
        int currentYear = DateTime.Now.Year;

        for (int year = currentYear - 5; year <= currentYear + 5; year++)
        {
            YearPicker.Items.Add(year.ToString());
        }

        for (int month = 1; month <= 12; month++)
        {
            MonthPicker.Items.Add(month.ToString());
        }

        YearPicker.SelectedItem = currentYear.ToString();
        MonthPicker.SelectedItem = DateTime.Now.Month.ToString();
    }

    /// <summary>
    /// 강의목표 추가 버튼 클릭 시
    /// </summary>
    /// <param name="sender"></param>
    /// <param name="e"></param>
    private void OnAddLessonObjectiveClicked(object sender, EventArgs e)
    {
        if (_lessonObjectiveEntries.Count >= MaxObjectives)
        {
            DisplayAlert("알림", $"강의목표는 최대 {MaxObjectives}개까지 입력할 수 있어요.", "확인");
            return;
        }

        Entry entry = new Entry
        {
            Placeholder = $"강의목표 {_lessonObjectiveEntries.Count + 1}"
        };

        _lessonObjectiveEntries.Add(entry);
        LessonObjectivesLayout.Children.Add(entry);
    }

    /// <summary>
    /// 강의계획 추가 버튼 클릭 시
    /// </summary>
    /// <param name="sender"></param>
    /// <param name="e"></param>
    private void OnAddLessonPlanClicked(object sender, EventArgs e)
    {
        if (_lessonPlanEntries.Count >= MaxLessonPlans)
        {
            DisplayAlert("알림", $"강의계획은 최대 {MaxLessonPlans}개까지 추가할 수 있어요.", "확인");
            return;
        }

        DatePicker datePicker = new()
        {
            Format = "yyyy-MM-dd",
            WidthRequest = 130
        };

        Entry planEntry = new()
        {
            Placeholder = "강의계획 내용",
            WidthRequest = 200
        };

        Entry remarkEntry = new()
        {
            Placeholder = "비고",
            WidthRequest = 200
        };

        HorizontalStackLayout layout = new()
        {
            Spacing = 10,
            Children = { datePicker, planEntry, remarkEntry }
        };

        _lessonPlanEntries.Add((datePicker, planEntry, remarkEntry));
        LessonPlansLayout.Children.Add(layout);
    }

    /// <summary>
    /// 강의계획서 생성 버튼 클릭 시
    /// </summary>
    /// <param name="sender"></param>
    /// <param name="e"></param>
    private async void OnGenerateLessonPlanClicked(object sender, EventArgs e)
    {
        try
        {
            // 엑셀 파일 저장 경로 설정 (플랫폼별로 다르게 처리해야 함)
            string fileName = $"LessonPlan_{DateTime.Now:yyyyMMddHHmmss}.xlsx";
            string filePath;

#if ANDROID
            // 안드로이드 다운로드 폴더 (권한 필요할 수 있음)
            var downloadsPath = Android.OS.Environment.GetExternalStoragePublicDirectory(Android.OS.Environment.DirectoryDownloads).AbsolutePath;
            filePath = System.IO.Path.Combine(downloadsPath, fileName);
#elif WINDOWS
        // 윈도우 데스크탑 내 문서 폴더
        var documentsPath = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
        filePath = System.IO.Path.Combine(documentsPath, fileName);
#else
        // 기타 플랫폼 적당한 임시 폴더
        filePath = System.IO.Path.Combine(FileSystem.CacheDirectory, fileName);
#endif

            using var package = new ExcelPackage(new FileInfo(filePath));
            var worksheet = package.Workbook.Worksheets.Add("강의계획서");

            // 헤더 작성
            worksheet.Cells[1, 1].Value = "강의일자";
            worksheet.Cells[1, 2].Value = "강의계획";
            worksheet.Cells[1, 3].Value = "비고";

            // 데이터 채우기
            int row = 2;
            foreach (var plan in LessonPlans)
            {
                worksheet.Cells[row, 1].Value = plan.LessonDate.ToString("yyyy-MM-dd");
                worksheet.Cells[row, 2].Value = plan.PlanDetail;
                worksheet.Cells[row, 3].Value = plan.Note;
                row++;
            }

            // 컬럼 너비 자동 조정
            worksheet.Cells[worksheet.Dimension.Address].AutoFitColumns();

            // 저장
            await package.SaveAsync();

            // 사용자에게 완료 메시지 띄우기
            await DisplayAlert("완료", $"강의계획서가 생성되었습니다.\n파일 위치:\n{filePath}", "확인");

        }
        catch (Exception ex)
        {
            await DisplayAlert("오류", $"엑셀 생성 중 오류가 발생했습니다.\n{ex.Message}", "확인");
        }
    }

}
