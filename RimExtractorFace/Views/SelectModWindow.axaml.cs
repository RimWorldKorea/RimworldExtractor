using Avalonia.Controls;
using Avalonia.Input;
using MsBox.Avalonia;
using RimExtractorCore.DataTypes;
using RimExtractorFace.Services;
using RimExtractorFace.ViewModels;

namespace RimExtractorFace.Views;

public partial class SelectModWindow : Window
{
    private readonly SelectModViewModel _viewModel;

    public Mod? SelectedMod => _viewModel.SelectedMod;
    public List<Mod> ReferenceMods => _viewModel.ReferenceMods;

    public bool IsSuccess { get; private set; } = false;

    public SelectModWindow() : this(null) { }

    public SelectModWindow(Mod? initialMod = null, List<Mod>? initialReferenceMods = null)
    {
        InitializeComponent();
        
        // 의존성 주입하여 뷰모델 생성
        _viewModel = new SelectModViewModel(
            new ExternalProcessService(),
            new ExtractionService(),
            initialMod,
            initialReferenceMods);

        DataContext = _viewModel;

        _viewModel.RequestShowAlert += async (msg) =>
        {
            var box = MessageBoxManager.GetMessageBoxStandard("알림", msg);
            await box.ShowWindowDialogAsync(this);
        };

        _viewModel.RequestClose += () =>
        {
            IsSuccess = true;
            Close();
        };
        
        Loaded += async (s, e) => await _viewModel.InitializeAsync();
    }
    
    private void OnItemPointerPressed(object? sender, PointerPressedEventArgs e)
    {
        if (e.GetCurrentPoint(this).Properties.IsRightButtonPressed)
        {
            // 우클릭 이벤트가 ListBox로 버블링되어 아이템이 선택되는 것을 차단합니다.
            e.Handled = true; 
            
            // 컨텍스트 메뉴를 수동으로 강제 호출합니다.
            if (sender is Control control && control.ContextMenu != null)
                control.ContextMenu.Open(control); 
        }
    }
}