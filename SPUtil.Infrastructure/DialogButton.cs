using System;
using System.Windows.Input; // Обязательно для ICommand
using Prism.Commands;



public class DialogButton
{
	public required string Caption { get; set; }
	public required Action Action { get; set; }
    public bool IsCancel { get; set; }

    // Добавьте это свойство, чтобы XAML мог "нажать" на кнопку
    public ICommand ClickCommand => new DelegateCommand(() => Action?.Invoke());
}