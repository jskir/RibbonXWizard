using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Windows;
using System.Windows.Input;
using Microsoft.Win32;
using RibbonXWizard.Data;
using RibbonXWizard.Models;
using RibbonXWizard.Services;

namespace RibbonXWizard.ViewModels
{
    /// <summary>
    /// ViewModel for the main window - handles UI logic and user interactions
    /// </summary>
    public class MainViewModel : INotifyPropertyChanged
    {
        private readonly RibbonXmlGenerator _xmlGenerator;
        private readonly DotmProcessor _dotmProcessor;

        // Backing fields for properties
        private BuiltInTabs.TabInfo? _selectedTab;
        private string _groupLabel = "My Tools";
        private string _buttonLabel = "My Button";
        private string _screenTip = "";
        private string _imageMso = "MacroPlay";
        private string _macroName = "";
        private string _sourceDotmPath = "";
        private string _generatedXml = "";
        private string _generatedVba = "";
        private string _instructions = "";
        private bool _hasGenerated = false;

        public MainViewModel()
        {
            _xmlGenerator = new RibbonXmlGenerator();
            _dotmProcessor = new DotmProcessor();

            // Initialize available tabs
            AvailableTabs = BuiltInTabs.GetWordTabs();
            SelectedTab = AvailableTabs.FirstOrDefault();

            // Initialize commands
            BrowseSourceFileCommand = new RelayCommand(BrowseSourceFile);
            GenerateCommand = new RelayCommand(Generate, CanGenerate);
            SaveOutputCommand = new RelayCommand(SaveOutput, CanSaveOutput);
            CopyVbaCommand = new RelayCommand(CopyVba, CanCopyVba);
            CopyInstructionsCommand = new RelayCommand(CopyInstructions, CanCopyInstructions);
        }

        #region Properties

        public List<BuiltInTabs.TabInfo> AvailableTabs { get; }

        public BuiltInTabs.TabInfo? SelectedTab
        {
            get => _selectedTab;
            set { _selectedTab = value; OnPropertyChanged(); }
        }

        public string GroupLabel
        {
            get => _groupLabel;
            set { _groupLabel = value; OnPropertyChanged(); }
        }

        public string ButtonLabel
        {
            get => _buttonLabel;
            set { _buttonLabel = value; OnPropertyChanged(); }
        }

        public string ScreenTip
        {
            get => _screenTip;
            set { _screenTip = value; OnPropertyChanged(); }
        }

        public string ImageMso
        {
            get => _imageMso;
            set { _imageMso = value; OnPropertyChanged(); }
        }

        public string MacroName
        {
            get => _macroName;
            set { _macroName = value; OnPropertyChanged(); }
        }

        public string SourceDotmPath
        {
            get => _sourceDotmPath;
            set { _sourceDotmPath = value; OnPropertyChanged(); }
        }

        public string GeneratedXml
        {
            get => _generatedXml;
            set { _generatedXml = value; OnPropertyChanged(); }
        }

        public string GeneratedVba
        {
            get => _generatedVba;
            set { _generatedVba = value; OnPropertyChanged(); }
        }

        public string Instructions
        {
            get => _instructions;
            set { _instructions = value; OnPropertyChanged(); }
        }

        public bool HasGenerated
        {
            get => _hasGenerated;
            set { _hasGenerated = value; OnPropertyChanged(); }
        }

        #endregion

        #region Commands

        public ICommand BrowseSourceFileCommand { get; }
        public ICommand GenerateCommand { get; }
        public ICommand SaveOutputCommand { get; }
        public ICommand CopyVbaCommand { get; }
        public ICommand CopyInstructionsCommand { get; }

        private void BrowseSourceFile()
        {
            var dialog = new OpenFileDialog
            {
                Filter = "Word Macro-Enabled Templates (*.dotm)|*.dotm|All Files (*.*)|*.*",
                Title = "Select Source .dotm File"
            };

            if (dialog.ShowDialog() == true)
            {
                SourceDotmPath = dialog.FileName;
            }
        }

        private bool CanGenerate()
        {
            return !string.IsNullOrWhiteSpace(SourceDotmPath) &&
                   !string.IsNullOrWhiteSpace(ButtonLabel) &&
                   !string.IsNullOrWhiteSpace(MacroName) &&
                   File.Exists(SourceDotmPath);
        }

        private void Generate()
        {
            try
            {
                // Create configuration from UI inputs
                var config = new RibbonConfig
                {
                    TargetTabId = SelectedTab?.IdMso ?? "TabHome",
                    GroupLabel = GroupLabel,
                    ButtonLabel = ButtonLabel,
                    ScreenTip = ScreenTip,
                    ImageMso = ImageMso,
                    MacroName = MacroName,
                    SourceDotmPath = SourceDotmPath
                };

                // Generate the ribbon customization
                var result = _xmlGenerator.Generate(config);

                // Update UI with results
                GeneratedXml = result.XmlContent;
                GeneratedVba = result.VbaCallbackCode;
                Instructions = result.Instructions;
                HasGenerated = true;

                MessageBox.Show(
                    "Ribbon customization generated successfully!\n\n" +
                    "Click 'Save Output File' to create your customized .dotm file.",
                    "Success",
                    MessageBoxButton.OK,
                    MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    $"Error generating ribbon customization:\n\n{ex.Message}",
                    "Error",
                    MessageBoxButton.OK,
                    MessageBoxImage.Error);
            }
        }

        private bool CanSaveOutput()
        {
            return HasGenerated && !string.IsNullOrWhiteSpace(GeneratedXml);
        }

        private void SaveOutput()
        {
            try
            {
                var dialog = new SaveFileDialog
                {
                    Filter = "Word Macro-Enabled Templates (*.dotm)|*.dotm",
                    Title = "Save Customized .dotm File",
                    FileName = Path.GetFileNameWithoutExtension(SourceDotmPath) + "_CustomRibbon.dotm"
                };

                if (dialog.ShowDialog() == true)
                {
                    // Process the file
                    _dotmProcessor.ApplyCustomization(
                        SourceDotmPath,
                        dialog.FileName,
                        GeneratedXml);

                    MessageBox.Show(
                        $"File saved successfully!\n\n{dialog.FileName}\n\n" +
                        "Don't forget to add the VBA callback code (see instructions below).",
                        "Success",
                        MessageBoxButton.OK,
                        MessageBoxImage.Information);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    $"Error saving file:\n\n{ex.Message}",
                    "Error",
                    MessageBoxButton.OK,
                    MessageBoxImage.Error);
            }
        }

        private bool CanCopyVba()
        {
            return !string.IsNullOrWhiteSpace(GeneratedVba);
        }

        private void CopyVba()
        {
            try
            {
                Clipboard.SetText(GeneratedVba);
                MessageBox.Show("VBA code copied to clipboard!", "Success",
                    MessageBoxButton.OK, MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error copying to clipboard:\n\n{ex.Message}",
                    "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private bool CanCopyInstructions()
        {
            return !string.IsNullOrWhiteSpace(Instructions);
        }

        private void CopyInstructions()
        {
            try
            {
                Clipboard.SetText(Instructions);
                MessageBox.Show("Instructions copied to clipboard!", "Success",
                    MessageBoxButton.OK, MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error copying to clipboard:\n\n{ex.Message}",
                    "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        #endregion

        #region INotifyPropertyChanged

        public event PropertyChangedEventHandler? PropertyChanged;

        protected virtual void OnPropertyChanged([CallerMemberName] string? propertyName = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }

        #endregion

        #region RelayCommand Helper

        /// <summary>
        /// Simple ICommand implementation for WPF commands
        /// </summary>
        private class RelayCommand : ICommand
        {
            private readonly Action _execute;
            private readonly Func<bool>? _canExecute;

            public RelayCommand(Action execute, Func<bool>? canExecute = null)
            {
                _execute = execute ?? throw new ArgumentNullException(nameof(execute));
                _canExecute = canExecute;
            }

            public event EventHandler? CanExecuteChanged
            {
                add => CommandManager.RequerySuggested += value;
                remove => CommandManager.RequerySuggested -= value;
            }

            public bool CanExecute(object? parameter) => _canExecute?.Invoke() ?? true;

            public void Execute(object? parameter) => _execute();
        }

        #endregion
    }
}