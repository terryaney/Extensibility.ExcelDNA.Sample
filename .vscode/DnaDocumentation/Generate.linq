<Query Kind="Program">
  <Reference Relative="..\..\src\bin\Debug\net7.0-windows\KAT.Camelot.Extensibility.Excel.AddIn.dll">C:\BTR\Camelot\Extensibility\Excel.AddIn\src\bin\Debug\net7.0-windows\KAT.Camelot.Extensibility.Excel.AddIn.dll</Reference>
  <NuGetReference>CliWrap</NuGetReference>
  <Namespace>CliWrap</Namespace>
  <Namespace>CliWrap.Buffered</Namespace>
  <Namespace>ExcelDna.Integration</Namespace>
  <Namespace>KAT.Camelot.Extensibility.Excel.AddIn</Namespace>
  <Namespace>System.Threading.Tasks</Namespace>
  <RuntimeVersion>7.0</RuntimeVersion>
</Query>

async Task Main( string[] args )
{
	try
	{
		Func<string, string, string?> getArgument = (a, key) =>
		{
			if (a.StartsWith(key + ":", StringComparison.InvariantCultureIgnoreCase))
			{
				return a.Substring(a.IndexOf(":") + 1);
			}

			return null;
		};

		formatOutput = args.Select(a => getArgument(a, "tool")).FirstOrDefault(a => !string.IsNullOrEmpty(a)) != "Visual Studio";
		var isDebug = args.Select(a => getArgument(a, "configurationname")).FirstOrDefault(a => !string.IsNullOrEmpty(a)) == "Debug";

		WriteLine($"\r\n========== KAT Tools Documentation Processing ==========\r\n", false);

		var assembly = typeof(Ribbon).Assembly;

		WriteLine("Detecting DNA Functions...", true);
		var info =
			assembly.GetTypes()
				.SelectMany(t => t.GetMethods())
				.Where(m => (m.GetCustomAttribute<KatExcelFunctionAttribute>() ?? m.GetCustomAttribute<ExcelFunctionAttribute>()) != null)
				.Select(m =>
				{
					var katFunc = m.GetCustomAttribute<KatExcelFunctionAttribute>();
					var dnaFunc = (katFunc as ExcelFunctionAttribute) ?? m.GetCustomAttribute<ExcelFunctionAttribute>()!;

					return new
					{
						Name = dnaFunc.Name ?? m.Name,
						Category = dnaFunc.Category,
						Description = katFunc?.Summary ?? dnaFunc.Description,
						Returns = katFunc?.Returns,
						Remarks = katFunc?.Remarks,
						Example = katFunc?.Example,
						HelpTopic = dnaFunc.HelpTopic,
						CreateDebugFunction = katFunc?.CreateDebugFunction ?? false,
						Arguments =
							m.GetParameters()
								.Select(p =>
								{
									var katArg = p.GetCustomAttribute<KatExcelArgumentAttribute>();
									var dnaArg = (katArg as ExcelArgumentAttribute) ?? p.GetCustomAttribute<ExcelArgumentAttribute>();
									return new
									{
										Name = katArg?.DisplayName ?? dnaArg?.Name ?? p.Name!,
										Description = katArg?.Summary ?? dnaArg?.Description ?? "TODO: Document this parameter.",
										Type = katArg?.Type ?? p.ParameterType,
										IsOptional = p.IsOptional,
										DefaultValue = katArg?.Default ?? p.DefaultValue?.ToString()
									};
								})
					};
				})
				.ToArray();

		XNamespace ns = "http://schemas.excel-dna.net/intellisense/1.0";

		var fileName = isDebug ? "KAT.Extensibility.Excel.Debug.intellisense.xml" : "KAT.Extensibility.Excel.intellisense.xml";
		WriteLine($"Generating {fileName}...", true);
		var intelliSense =
			new XElement(ns + "IntelliSense",
				new XElement(ns + "FunctionInfo",
					info.Select(i =>
						new XElement(ns + "Function",
							new XAttribute("Name", i.Name),
							new XAttribute("Description", i.Description),
							i.HelpTopic != null ? new XAttribute("HelpTopic", i.HelpTopic) : null,
							i.Arguments.Select(a =>
								new XElement(ns + "Argument",
									new XAttribute("Name", a.Name),
									new XAttribute("Description", a.Description)
								)
							)
						)
					)
				)
			);

		WriteLine($"Saving {fileName}", true);
		if ( isDebug )
		{
			intelliSense.Save(@$"C:\BTR\Camelot\Extensibility\Excel.AddIn\src\bin\Debug\net7.0-windows\{fileName}");
		}
		else
		{
			intelliSense.Save(@$"C:\BTR\Camelot\Extensibility\Excel.AddIn\.vscode\DnaDocumentation\{fileName}");
			intelliSense.Save(@$"C:\BTR\Camelot\Extensibility\Excel.AddIn\dist\Resources\{fileName}");
		}

		WriteLine("Generating Markdown Documentation...", true);
		var functionCategories = info
			.OrderBy(i => i.Name)
			.GroupBy(g => g.Category);

		var validFiles = new List<string>();
		validFiles.Add("RBLe.md");

		foreach (var category in functionCategories)
		{
			var categoryName = category.Key.Replace(" ", "");
			var templateMd = @$"C:\BTR\Camelot\Extensibility\Excel.AddIn\.vscode\DnaDocumentation\{categoryName}.md";
			var templateContent = await File.ReadAllTextAsync(templateMd);

			var categoryTable = new StringBuilder();
			foreach (var m in category)
			{
				categoryTable.AppendLine($"[`{m.Name}`](RBLe{categoryName}.{m.Name}.md) | {m.Description}");

				if (m.CreateDebugFunction)
				{
					categoryTable.AppendLine($"[`{m.Name}Debug`](RBLe{categoryName}.{m.Name}Debug.md) | Debug version of {m.Name} that returns value or exception string (instead of #VALUE on error).  {m.Description}");
				}

				var returns = m.Returns != null
					? $"{Environment.NewLine}**Returns:** {m.Returns}"
					: null;

				var remarks = m.Remarks != null
					? $"{Environment.NewLine}## Remarks{Environment.NewLine + Environment.NewLine}{string.Join("  " + Environment.NewLine, m.Remarks.Split(Environment.NewLine))}"
					: null;

				var example = m.Example != null
					? $"{Environment.NewLine}## Example{Environment.NewLine + Environment.NewLine}{m.Example}"
					: null;

				var parameters =
					string.Join(
						Environment.NewLine,
						m.Arguments.Select(a =>
						{
							var defaultValue = a.IsOptional ? a.DefaultValue : null;

							if (defaultValue != null && !defaultValue.StartsWith("`"))
							{
								defaultValue = a.Type == typeof(string)
									? $"`\"{defaultValue}\"`"
									: $"`{defaultValue}`";
							}

							return $"`{a.Name}` | {a.Type.Name + (a.IsOptional ? "?" : "")} | {defaultValue} | {a.Description.Replace("|", "\\|")}";
						})
					);

				var functionTemplate = @$"# {m.Name} Function

{m.Description}
{string.Join("", new[] { returns, remarks }.Where(i => i != null))}
## Syntax

```excel
={m.Name}({string.Join(", ", m.Arguments.Select(a => a.Name))})
```

Parameter | Type | Default | Description
---|---|---|---
{parameters}
{example}
[Back to {category.Key}](RBLe{categoryName}.md) | [Back to All RBLe Functions](RBLe.md#function-documentation)";

				var functionPath = @$"C:\BTR\Documentation\Camelot\RBLe\RBLe{categoryName}.{m.Name}.md";
				validFiles.Add(Path.GetFileName(functionPath));
				await File.WriteAllTextAsync(functionPath, functionTemplate);


				if (m.CreateDebugFunction)
				{
					functionPath = @$"C:\BTR\Documentation\Camelot\RBLe\RBLe{categoryName}.{m.Name}Debug.md";
					validFiles.Add(Path.GetFileName(functionPath));
					await File.WriteAllTextAsync(
						functionPath,
						functionTemplate
							.Replace(m.Name, m.Name + "Debug")
							.Replace(m.Description, $"Debug version of {m.Name} that returns value or exception string (instead of #VALUE on error).  {m.Description}")
					);
				}
			}

			var categoryPath = @$"C:\BTR\Documentation\Camelot\RBLe\RBLe{categoryName}.md";
			validFiles.Add(Path.GetFileName(categoryPath));
			await File.WriteAllTextAsync(
				categoryPath,
				templateContent.Replace("{FUNCTIONS}", categoryTable.ToString())
			);
		}

		var filesToDelete = new DirectoryInfo(@"C:\BTR\Documentation\Camelot\RBLe\").GetFiles("RBLe*.*").Where(f => !validFiles.Contains(f.Name));
		foreach (var f in filesToDelete)
		{
			File.Delete(f.FullName);
		}

		var gitPath = @"C:\BTR\Documentation\Camelot";
		var currentBuildBranch = (await GetRepositoryBranchesAsync(gitPath)).Single(b => b.IsActive);

		if (currentBuildBranch.NeedsCommit && currentBuildBranch.StatusLog.Any(s => s.Contains("RBLe/RBLe")))
		{
			WriteLine("Committing Markdown Documentation...", true);
			await CallGitCommandLineAsync(gitPath, new[] { "add", "RBLe/RBLe*" });
			await CallGitCommandLineAsync(gitPath, new[] { "com", "-m", $"RBLe Function Documentation" });
			WriteLine("Pushing Markdown Documentation...", true);
			await CallGitCommandLineAsync(gitPath, new[] { "push" });
		}

#if !CMD
		functionCategories.Dump();
#endif

		WriteLine($"\r\n========== KAT Tools Documentation Complete ==========\r\n", true);
		Environment.Exit(Environment.ExitCode = 0);
	}
	catch (Exception ex)
	{
		WriteLine($"\r\n========== KAT Tools Documentation Processing: \x1b[31mfailed\x1b[0m ==========\r\n");
		WriteLine("\x1b[31m");
		WriteLine($"{ex.GetType().ToString()}: {ex.Message}");
		var inner = ex.InnerException;
		if (inner != null)
		{
			WriteLine($"{Environment.NewLine}Inner: {inner.GetType().ToString()}: {inner.Message}");

		}
		WriteLine($"{Environment.NewLine}{(inner ?? ex).StackTrace}");
		WriteLine("\x1b[0m");
		Environment.Exit(Environment.ExitCode = -1);
	}
}

public static class Extensions
{
	public static bool IsNullable(this Type type) => Nullable.GetUnderlyingType( type ) != null;
}

Regex branchRegEx = new Regex(@"^(?<branch>\S*)\s*(?<commit>\S*)\s*(?:\[ahead (?<ahead>\d+)(?:, behind (?<behind>\d+))?\]|\[behind (?<behindOnly>\d+)\])?\s*(?<comment>.*)", RegexOptions.Compiled);

async Task<GitBranch[]> GetRepositoryBranchesAsync(string repositoryPath)
{
	var statusRaw = await CallGitCommandLineAsync(repositoryPath, new[] { "st" });
	var branchesRaw = await CallGitCommandLineAsync(repositoryPath, new[] { "br", "-v" });

	return (
		from b in branchesRaw

		let match = branchRegEx.Match(b.Substring(2))
		let branch = match.Groups["branch"].Value
		let branchParts = branch.Split('/')
		let isRemote = string.Compare(branchParts[0], "remotes", true) == 0

		let branchName = isRemote
			? string.Join("/", branchParts.Skip(2))
			: string.Join("/", branchParts)
		let remote = isRemote
			? branchParts[1]
			: null

		let ahead = match.Groups["ahead"].Success ? int.Parse(match.Groups["ahead"].Value) : 0
		let behindGroup = match.Groups["behind"].Success ? match.Groups["behind"] : match.Groups["behindOnly"]
		let behind = behindGroup.Success ? int.Parse(behindGroup.Value) : 0
		let needsCommit = b.StartsWith("*") && statusRaw.Length > 1

		where string.Compare(branchName, "HEAD", true) != 0

		select new GitBranch
		{
			Branch = branchName,
			Remote = remote ?? "origin",
			Commit = match.Groups["commit"].Value,
			Comment = match.Groups["comment"].Value,
			IsRemote = isRemote,
			IsActive = b.StartsWith("*"),
			NeedsCommit = needsCommit,
			NeedsSync = b.StartsWith("*") && !needsCommit && (ahead + behind) > 0,
			Ahead = ahead,
			Behind = behind,
			StatusLog = statusRaw
			// FeatureBranch = ( string.Compare( branchName, "main", true ) != 0 && string.Compare( branchName, "_Test", true ) != 0 )
		}
	).ToArray();
}

class GitBranch
{
	public required string Branch { get; set; }
	public required string Remote { get; set; }
	public required string Commit { get; set; }
	public required string Comment { get; set; }
	public bool IsRemote { get; set; }
	public bool IsActive { get; set; }
	public bool NeedsCommit { get; set; }
	public bool NeedsSync { get; set; }
	public int Ahead { get; set; }
	public int Behind { get; set; }
	public string[] StatusLog { get; set; } = Array.Empty<string>();
}

async Task<string[]> CallGitCommandLineAsync(string repoPath, string[] arguments)
{
	// http://stackoverflow.com/questions/6127063/running-git-from-windows-cmd-line-where-are-key-files
	// SSH requires HOME environment variable to be set.

	try
	{
		var results =
			await Cli.Wrap(@"C:\Program Files\Git\bin\git.exe")
				.WithWorkingDirectory(repoPath)
				.WithArguments(arguments)
				.ExecuteBufferedAsync();

		var logRaw = string.IsNullOrEmpty(results.StandardOutput) && !string.IsNullOrEmpty(results.StandardError)
			? results.StandardError.Split('\n').ToArray()
			: results.StandardOutput.Split('\n').ToArray();

		return logRaw.Where(b => !string.IsNullOrEmpty(b)).ToArray();
	}
	catch (Exception ex)
	{
		throw new ApplicationException($"Unable to issue 'git {string.Join(" ", arguments)}' to {repoPath}", ex);
	}
}

StringBuilder? output = null;
bool formatOutput = true;
void WriteLine(string line, bool flush = true)
{
	var l =
		line.IndexOf("FAIL") > -1 ? $"\x1b[31m{line}\x1b[0m" :
		line.IndexOf("WARN") > -1 ? $"\x1b[33m{line}\x1b[0m" :
		line;

	if (!formatOutput)
	{
		// For Visual Studio, install VSColorOutput extension and add .* failed .* regular expression as the first expression it processes
		l = l.Replace("\x1b[31m", "").Replace("\x1b[33m", "").Replace("\x1b[0m", "");
	}

	if (!flush)
	{
		if (output == null)
		{
			output = new StringBuilder();
		}
		output.AppendLine(l);
	}
	else
	{
		if (output != null)
		{
			Console.WriteLine(output.ToString());
			output = null;
		}
		Console.WriteLine(l);
	}
}