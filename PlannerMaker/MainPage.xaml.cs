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
        // TODO: 여기서 엑셀 생성 로직 호출
        // 플랫폼별 저장/다운로드 구현 예정

        // 임시로 알림 띄우기
        await DisplayAlert("알림", "강의계획서 생성 기능을 곧 구현할게요!", "확인");
    }

}
