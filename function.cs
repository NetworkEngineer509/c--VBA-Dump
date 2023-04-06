using System;
using System.IO;
using System.Text;

class Program
{
    static void Main(string[] args)
    {
        string binFile = "vbaProject.bin";
        string dumpDirectory = "VBA Dump";

        // Load the vbaProject.bin file into a byte array
        byte[] fileBytes = File.ReadAllBytes(binFile);

        // Get the starting offset of the VBA project data
        int projectOffset = BitConverter.ToInt32(fileBytes, 48);

        // Get the number of projects in the file
        int projectCount = BitConverter.ToInt16(fileBytes, projectOffset + 6);

        // Find the start of the VBA project data for the first project
        int vbProjectOffset = projectOffset + BitConverter.ToInt32(fileBytes, projectOffset + 16);

        // Loop through all projects in the file
        for (int i = 0; i < projectCount; i++)
        {
            // Find the start of the project stream data
            int projectStreamOffset = vbProjectOffset + 28 + BitConverter.ToInt32(fileBytes, vbProjectOffset + 24);

            // Get the number of modules in the project
            int moduleCount = BitConverter.ToInt32(fileBytes, projectStreamOffset + 8);

            // Find the start of the module data
            int moduleOffset = projectStreamOffset + BitConverter.ToInt32(fileBytes, projectStreamOffset + 12);

            // Loop through all modules in the project
            for (int j = 0; j < moduleCount; j++)
            {
                // Get the name and code of the module
                int nameOffset = moduleOffset + BitConverter.ToInt32(fileBytes, moduleOffset + 8);
                int codeOffset = moduleOffset + BitConverter.ToInt32(fileBytes, moduleOffset + 20);
                int codeLength = BitConverter.ToInt32(fileBytes, moduleOffset + 24);
                string moduleName = Encoding.Unicode.GetString(fileBytes, nameOffset, codeOffset - nameOffset - 2);
                string moduleCode = Encoding.Unicode.GetString(fileBytes, codeOffset, codeLength);

                // Create a new file in the dump directory for the module
                string fileName = $"{moduleName}.bas";
                string filePath = Path.Combine(dumpDirectory, fileName);
                File.WriteAllText(filePath, moduleCode);

                // Advance to the next module
                moduleOffset += 28;
            }

            // Advance to the next project
            vbProjectOffset += 64;
        }

        Console.WriteLine($"Dumped {projectCount} projects to {dumpDirectory}");
    }
}
