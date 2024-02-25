using System.Diagnostics;
using System.Reflection;

namespace RimworldExtractorGUI
{
    /*
    * TODO:
    * Patches �ڵ� ���� ��� ����
    */


    /*
     * �۾� ����:
     * Internal)
     * 1. Translation Handle �±��� �տ� '*'�� ������ �̴� Type Ÿ������ �����Ͽ�, �븻����¡�� �ٸ��� �����ϵ��� ���� (���� Ż���� MVCF.Comp_VerbProps ���� ����)
     * 2. PatchOperationAdd ���� ��, xpath�� �ִ� Def�� ���� ����� Def�̸� ������ ���ϴ� ���� �ذ� (���� �丣�ҳ� ���� ���� ����)
     * 3. Common ���� ���
     * 4. XMLExtension�� ������ Patches�� ���� ��쵵 ��� ([FSF] FrozenSnowFox Tweaks ���� ����)
     * 5. Patches �ڵ� ���� ��� ���� (�Ź� FindMod -> Replace�� �ƴ϶�, FindMod�� ���� �ֵ鳢���� �ϳ��� Sequence���� �����ϵ���)
     * 6. ������ �⺻���� Languages XML�� ���� (���� ����� �⺻���� �����ϰ� �ϱ� ����)
     *
     * GUI)
     * 1. UI ����
     * 2. 
     */
    internal static class Program
    {
        /// <summary>
        ///  The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main()
        {
            AppDomain.CurrentDomain.AssemblyResolve += CurrentDomainOnAssemblyResolve;
            // To customize application configuration such as set high DPI settings or default font,
            // see https://aka.ms/applicationconfiguration.
            ApplicationConfiguration.Initialize();
            if (!File.Exists("Prefabs.dat"))
            {
                var formInitialPathSelect = new FormInitialPathSelect();
                formInitialPathSelect.StartPosition = FormStartPosition.CenterScreen;
                if (formInitialPathSelect.ShowDialog() != DialogResult.OK)
                {
                    MessageBox.Show("���� ������ �Ϸ����ּ���.");
                    return;
                }
                // Application.Run();
            }

            var formMain = new FormMain();
            formMain.StartPosition = FormStartPosition.CenterScreen;
            Application.Run(formMain);
        }

        private static Assembly? CurrentDomainOnAssemblyResolve(object? sender, ResolveEventArgs args)
        {
            // Ignore missing resources
            if (args.Name.Contains(".resources"))
                return null;

            // check for assemblies already loaded
            Assembly? assembly = AppDomain.CurrentDomain.GetAssemblies().FirstOrDefault(a => a.FullName == args.Name);
            if (assembly != null)
                return assembly;

            string filename = args.Name.Split(',')[0] + ".dll".ToLower();
            var assemblyFilePath = Path.Combine("bin", filename);

            if (File.Exists(assemblyFilePath))
            {
                try
                {
                    return Assembly.LoadFrom(assemblyFilePath);
                }
                catch
                {
                    return null;
                }
            }
            return null;
        }
    }
}