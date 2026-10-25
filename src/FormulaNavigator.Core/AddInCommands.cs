using System;
using System.Collections.Generic;
using System.Linq;

namespace FormulaNavigator.Core
{
    // The complete set of user-facing commands and their shortcut keys. Exactly two
    // commands are exposed; the shared test harness asserts that the add-in
    // exports no others.
    public static class AddInCommands
    {
        // Excel-DNA shortcut syntax: "^q" is Ctrl+Q, "^Q" (capital) is Ctrl+Shift+Q.
        public const string ExploreKey = "^q";
        public const string DependentsKey = "^Q";
        public const string ExploreCommand = "FormulaNavigator_Explore";
        public const string DependentsCommand = "FormulaNavigator_Dependents";

        public static IReadOnlyList<CommandKeyBinding> Bindings
        {
            get
            {
                return new[]
                {
                    new CommandKeyBinding(ExploreKey, ExploreCommand),
                    new CommandKeyBinding(DependentsKey, DependentsCommand)
                }.ToList().AsReadOnly();
            }
        }
    }

    public sealed class CommandKeyBinding
    {
        public CommandKeyBinding(string key, string command)
        {
            if (String.IsNullOrWhiteSpace(key)) throw new ArgumentException("A key is required.", "key");
            if (String.IsNullOrWhiteSpace(command)) throw new ArgumentException("A command is required.", "command");
            Key = key;
            Command = command;
        }

        public string Key { get; private set; }
        public string Command { get; private set; }
    }
}
