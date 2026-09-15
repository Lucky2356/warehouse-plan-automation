using System.Windows;
using WarehousePlanAutomation.App.Infrastructure;
using WarehousePlanAutomation.Core.Abstractions;

namespace WarehousePlanAutomation.App.Views;

/// <summary>
/// Выбор одного варианта из тех, что нашлись в книге. Показывается посреди обработки,
/// когда правило допускает несколько ответов и угадывать нельзя.
/// </summary>
public partial class DecisionWindow : Window
{
    public DecisionWindow(DecisionRequest request)
    {
        InitializeComponent();

        QuestionText.Text = request.Question;
        ExplanationText.Text = request.Explanation;
        OptionsList.ItemsSource = request.Options;

        if (request.Options.Count > 0)
        {
            OptionsList.SelectedIndex = 0;
        }
    }

    /// <summary>Выбранный вариант или null, если человек нажал «Пропустить» или закрыл окно.</summary>
    public string? Choice { get; private set; }

    /// <summary>Заголовок красится в цвет темы - как у главного окна.</summary>
    protected override void OnSourceInitialized(EventArgs e)
    {
        base.OnSourceInitialized(e);
        TitleBarTheme.ApplyCurrent(this);
    }

    private void OnConfirm(object sender, RoutedEventArgs e)
    {
        Choice = OptionsList.SelectedItem as string;
        DialogResult = true;
    }

    private void OnSkip(object sender, RoutedEventArgs e)
    {
        Choice = null;
        DialogResult = false;
    }

    private void OnOptionDoubleClick(object sender, System.Windows.Input.MouseButtonEventArgs e)
    {
        if (OptionsList.SelectedItem is string)
        {
            OnConfirm(sender, e);
        }
    }
}
