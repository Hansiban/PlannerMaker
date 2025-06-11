using OfficeOpenXml;
using PlannerMaker.Models;
using System.Reflection;

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
        //유효성 검사

        try
        {
            // 현재 어셈블리
            var assembly = Assembly.GetExecutingAssembly();

            // 리소스 이름 확인 (이건 네임스페이스 + 폴더 + 파일명)
            string resourceName = "PlannerMaker.Resources.Templates.ExcelTemplate.xlsx";

            // 리소스 스트림 가져오기
            using Stream resourceStream = assembly.GetManifestResourceStream(resourceName);
            if (resourceStream == null)
            {
                throw new Exception($"리소스를 찾을 수 없습니다: {resourceName}");
            }

            // 패키지 생성
            using var package = new ExcelPackage(resourceStream);

            // 엑셀 수정 예시
            ExcelWorksheet worksheet = package.Workbook.Worksheets[0];

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
            //학원명, 강의일시, 담당과목
            worksheet.Cells["B4"].Value = $"{AcademyNameEntry.Text}";
            worksheet.Cells["B6"].Value = $"{YearPicker.SelectedItem.ToString()}년 {MonthPicker.SelectedItem.ToString()}월";
            worksheet.Cells["B7"].Value = $"{SubjectsEntry.Text}";

            //강의목표
            string strLessonObject = string.Empty;

            for (int i = 0; i < LessonObjectives.Count; i++)
            {
                string lessonObjectives = LessonObjectives[i];
                strLessonObject += $"◎ {lessonObjectives}";

                if (i != LessonObjectives.Count - 1)
                    strLessonObject += "\n\n";
            }

            worksheet.Cells["B8"].Value = strLessonObject;

            //강의계획
            for (int i = 0; i < MaxLessonPlans; i++)
            {
                LessonPlan lessonPlan = i >= LessonPlans.Count ? new LessonPlan() :  LessonPlans[i];

                string dateCellPoint = $"A{16 + ((i + 1) * 4)}";
                string detailCellPoint = $"B{16 + ((i + 1) * 4)}";
                string noteCellPoint = $"F{16 + ((i + 1) * 4)}";

                string strLessonDate = lessonPlan.LessonDate == DateTime.MinValue ? string.Empty : lessonPlan.LessonDate.ToString("MM/dd");

                worksheet.Cells[dateCellPoint].Value = $"{strLessonDate}";
                worksheet.Cells[detailCellPoint].Value = $"{lessonPlan.PlanDetail}";
                worksheet.Cells[noteCellPoint].Value = $"{lessonPlan.Note}";
            }

            // 컬럼 너비 자동 조정
            worksheet.Cells[worksheet.Dimension.Address].AutoFitColumns();

            // 저장
            var excelBytes = package.GetAsByteArray();
            File.WriteAllBytes(filePath, excelBytes);

            // 사용자에게 완료 메시지 띄우기
            await DisplayAlert("완료", $"강의계획서가 생성되었습니다.\n파일 위치:\n{filePath}", "확인");
        }
        catch (Exception ex)
        {
            await DisplayAlert("오류", $"엑셀 생성 중 오류가 발생했습니다.\n{ex.Message}", "확인");
        }
    }
}
