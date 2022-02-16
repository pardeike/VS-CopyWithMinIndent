using EnvDTE;
using Microsoft.VisualStudio.Shell;
using System;
using System.ComponentModel.Design;
using System.Linq;
using System.Text.RegularExpressions;
using System.Windows.Input;
using Task = System.Threading.Tasks.Task;

namespace CopyWithMinIndent
{
	internal sealed class CopySelectionCommand
	{
		public const int CommandId = 0x0100;
		public static readonly Guid CommandSet = new Guid("2cc1eee4-d7b3-4c83-8679-a7952afb4737");
		private readonly AsyncPackage package;

		private CopySelectionCommand(AsyncPackage package, OleMenuCommandService commandService)
		{
			this.package = package ?? throw new ArgumentNullException(nameof(package));
			commandService = commandService ?? throw new ArgumentNullException(nameof(commandService));

			var menuCommandID = new CommandID(CommandSet, CommandId);
			var menuItem = new MenuCommand(Execute, menuCommandID);
			commandService.AddCommand(menuItem);
		}

		public static CopySelectionCommand Instance
		{
			get;
			private set;
		}

		private IAsyncServiceProvider ServiceProvider => package;

		public static async Task InitializeAsync(AsyncPackage package)
		{
			await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync(package.DisposalToken);

			var commandService = await package.GetServiceAsync(typeof(IMenuCommandService)) as OleMenuCommandService;
			Instance = new CopySelectionCommand(package, commandService);
		}

		private void Execute(object sender, EventArgs e)
		{
			_ = ExecuteAsync(sender, e);
		}

		private static bool ModifierPressed()
		{
			return
				(Keyboard.GetKeyStates(Key.LeftShift) & KeyStates.Down) > 0 ||
				(Keyboard.GetKeyStates(Key.RightShift) & KeyStates.Down) > 0;
		}

		private async Task ExecuteAsync(object sender, EventArgs e)
		{
			await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync();

			var dte = await ServiceProvider.GetServiceAsync(typeof(DTE)) as DTE;
			var doc = dte?.ActiveDocument;
			if (doc != null)
			{
				var textDoc = doc.Object() as TextDocument;
				var selection = doc.Selection as TextSelection;
				var language = textDoc.Language;
				var rangeEnumerator = selection.TextRanges.GetEnumerator();
				var idx1 = selection.TopLine;
				var idx2 = selection.BottomLine;
				if (idx1 < idx2)
				{
					var text = textDoc.CreateEditPoint(textDoc.StartPoint).GetLines(idx1, idx2 + 1);
					var lines = Regex.Split(text, "\\r+\\n+").ToList();
					while (lines.Count > 0 && lines.First().Trim().Length == 0)
						lines.RemoveAt(0);
					while (lines.Count > 0 && lines.Last().Trim().Length == 0)
						lines.RemoveAt(lines.Count - 1);
					if (lines.Count > 0)
					{
						while (lines[0].Length > 0)
						{
							var c = lines[0].Substring(0, 1);
							if (lines.Any(line => line.StartsWith(c) == false))
								break;
							lines = lines.Select(line => line.Substring(1)).ToList();
						}
						var result = lines.Aggregate("", (prev, curr) => $"{prev}{(string.IsNullOrEmpty(prev) ? "" : "\n")}{curr}");
						if (ModifierPressed())
						{
							var type = language != "Plain Text" ? language : "";
							result = $"```{type}\n{result}\n```";
						}
						System.Windows.Forms.Clipboard.SetText(result);
					}
				}
			}
		}
	}
}
