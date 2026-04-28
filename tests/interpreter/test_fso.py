"""Tests for the VBScript interpreter."""

import pytest
import io
from pybasil import (
    Interpreter,
    parse,
    VBScriptError,
)


class TestFSOPathHelpers:
    """Test FileSystemObject path helper methods."""

    def test_buildpath(self):
        program = parse('''
        Dim fso
        Set fso = CreateObject("Scripting.FileSystemObject")
        result = fso.BuildPath("/tmp", "test.txt")
        ''')
        interp = Interpreter()
        interp.interpret(program)
        assert interp._environment.get('result') == '/tmp/test.txt'

    def test_getfilename(self):
        program = parse('''
        Dim fso
        Set fso = CreateObject("Scripting.FileSystemObject")
        result = fso.GetFileName("/path/to/file.txt")
        ''')
        interp = Interpreter()
        interp.interpret(program)
        assert interp._environment.get('result') == 'file.txt'

    def test_getbasename(self):
        program = parse('''
        Dim fso
        Set fso = CreateObject("Scripting.FileSystemObject")
        result = fso.GetBaseName("/path/to/file.txt")
        ''')
        interp = Interpreter()
        interp.interpret(program)
        assert interp._environment.get('result') == 'file'

    def test_getextensionname(self):
        program = parse('''
        Dim fso
        Set fso = CreateObject("Scripting.FileSystemObject")
        result = fso.GetExtensionName("/path/to/file.txt")
        ''')
        interp = Interpreter()
        interp.interpret(program)
        assert interp._environment.get('result') == 'txt'

    def test_getextensionname_no_ext(self):
        program = parse('''
        Dim fso
        Set fso = CreateObject("Scripting.FileSystemObject")
        result = fso.GetExtensionName("/path/to/file")
        ''')
        interp = Interpreter()
        interp.interpret(program)
        assert interp._environment.get('result') == ''

    def test_getparentfoldername(self):
        program = parse('''
        Dim fso
        Set fso = CreateObject("Scripting.FileSystemObject")
        result = fso.GetParentFolderName("/path/to/file.txt")
        ''')
        interp = Interpreter()
        interp.interpret(program)
        assert interp._environment.get('result') == '/path/to'

    def test_getabsolutepathname(self):
        program = parse('''
        Dim fso
        Set fso = CreateObject("Scripting.FileSystemObject")
        result = fso.GetAbsolutePathName(".")
        ''')
        interp = Interpreter()
        interp.interpret(program)
        import os
        assert interp._environment.get('result') == os.path.abspath('.')

    def test_gettempname(self):
        program = parse('''
        Dim fso
        Set fso = CreateObject("Scripting.FileSystemObject")
        result = fso.GetTempName
        ''')
        interp = Interpreter()
        interp.interpret(program)
        assert len(interp._environment.get('result')) > 0

class TestFSOFileExists:
    """Test FileSystemObject FileExists and FolderExists."""

    def test_fileexists_true(self, tmp_path):
        test_file = tmp_path / "exists.txt"
        test_file.write_text("hello")
        program = parse(f'Set fso = CreateObject("Scripting.FileSystemObject") : result = fso.FileExists("{test_file}")')
        interp = Interpreter()
        interp.interpret(program)
        assert interp._environment.get('result') is True

    def test_fileexists_false(self):
        program = parse('Set fso = CreateObject("Scripting.FileSystemObject") : result = fso.FileExists("/nonexistent_file_xyz")')
        interp = Interpreter()
        interp.interpret(program)
        assert interp._environment.get('result') is False

    def test_folderexists_true(self, tmp_path):
        program = parse(f'Set fso = CreateObject("Scripting.FileSystemObject") : result = fso.FolderExists("{tmp_path}")')
        interp = Interpreter()
        interp.interpret(program)
        assert interp._environment.get('result') is True

    def test_folderexists_false(self):
        program = parse('Set fso = CreateObject("Scripting.FileSystemObject") : result = fso.FolderExists("/nonexistent_dir_xyz")')
        interp = Interpreter()
        interp.interpret(program)
        assert interp._environment.get('result') is False

class TestFSOCreateAndWriteFile:
    """Test FileSystemObject CreateTextFile and writing."""

    def test_create_and_write(self, tmp_path):
        fpath = tmp_path / "test_write.txt"
        program = parse(f'''
        Dim fso, ts
        Set fso = CreateObject("Scripting.FileSystemObject")
        Set ts = fso.CreateTextFile("{fpath}")
        ts.Write "Hello"
        ts.Close
        ''')
        interp = Interpreter()
        interp.interpret(program)
        assert fpath.read_text() == "Hello"

    def test_writeline(self, tmp_path):
        fpath = tmp_path / "test_writeline.txt"
        program = parse(f'''
        Dim fso, ts
        Set fso = CreateObject("Scripting.FileSystemObject")
        Set ts = fso.CreateTextFile("{fpath}")
        ts.WriteLine "Line 1"
        ts.WriteLine "Line 2"
        ts.Close
        ''')
        interp = Interpreter()
        interp.interpret(program)
        content = fpath.read_text()
        assert "Line 1\n" in content
        assert "Line 2\n" in content

    def test_writeblanklines(self, tmp_path):
        fpath = tmp_path / "test_blank.txt"
        program = parse(f'''
        Dim fso, ts
        Set fso = CreateObject("Scripting.FileSystemObject")
        Set ts = fso.CreateTextFile("{fpath}")
        ts.Write "A"
        ts.WriteBlankLines 3
        ts.Write "B"
        ts.Close
        ''')
        interp = Interpreter()
        interp.interpret(program)
        content = fpath.read_text()
        assert content == "A\n\n\nB"

    def test_create_no_overwrite_error(self, tmp_path):
        fpath = tmp_path / "test_nooverwrite.txt"
        fpath.write_text("existing")
        program = parse(f'''
        Dim fso, ts
        Set fso = CreateObject("Scripting.FileSystemObject")
        Set ts = fso.CreateTextFile("{fpath}", False)
        ''')
        interp = Interpreter()
        with pytest.raises(VBScriptError, match="already exists"):
            interp.interpret(program)

class TestFSOOpenAndReadFile:
    """Test FileSystemObject OpenTextFile and reading."""

    def test_readall(self, tmp_path):
        fpath = tmp_path / "test_read.txt"
        fpath.write_text("Hello World")
        program = parse(f'''
        Dim fso, ts
        Set fso = CreateObject("Scripting.FileSystemObject")
        Set ts = fso.OpenTextFile("{fpath}")
        result = ts.ReadAll
        ts.Close
        ''')
        interp = Interpreter()
        interp.interpret(program)
        assert interp._environment.get('result') == "Hello World"

    def test_readline(self, tmp_path):
        fpath = tmp_path / "test_readline.txt"
        fpath.write_text("Line1\nLine2\nLine3\n")
        output = io.StringIO()
        program = parse(f'''
        Dim fso, ts
        Set fso = CreateObject("Scripting.FileSystemObject")
        Set ts = fso.OpenTextFile("{fpath}")
        Do While Not ts.AtEndOfStream
            WScript.Echo ts.ReadLine
        Loop
        ts.Close
        ''')
        interp = Interpreter(output_stream=output)
        interp.interpret(program)
        lines = output.getvalue().strip().split('\n')
        assert lines == ['Line1', 'Line2', 'Line3']

    def test_read_characters(self, tmp_path):
        fpath = tmp_path / "test_readchars.txt"
        fpath.write_text("Hello World")
        program = parse(f'''
        Dim fso, ts
        Set fso = CreateObject("Scripting.FileSystemObject")
        Set ts = fso.OpenTextFile("{fpath}")
        result = ts.Read(5)
        ts.Close
        ''')
        interp = Interpreter()
        interp.interpret(program)
        assert interp._environment.get('result') == "Hello"

    def test_open_nonexistent_error(self):
        program = parse('''
        Dim fso, ts
        Set fso = CreateObject("Scripting.FileSystemObject")
        Set ts = fso.OpenTextFile("/nonexistent_xyz.txt")
        ''')
        interp = Interpreter()
        with pytest.raises(VBScriptError, match="not found"):
            interp.interpret(program)

    def test_open_for_appending(self, tmp_path):
        fpath = tmp_path / "test_append.txt"
        fpath.write_text("Hello")
        program = parse(f'''
        Dim fso, ts
        Set fso = CreateObject("Scripting.FileSystemObject")
        Set ts = fso.OpenTextFile("{fpath}", 8)
        ts.Write " World"
        ts.Close
        ''')
        interp = Interpreter()
        interp.interpret(program)
        assert fpath.read_text() == "Hello World"

    def test_textstream_line_property(self, tmp_path):
        fpath = tmp_path / "test_line.txt"
        fpath.write_text("A\nB\nC\n")
        program = parse(f'''
        Dim fso, ts
        Set fso = CreateObject("Scripting.FileSystemObject")
        Set ts = fso.OpenTextFile("{fpath}")
        ts.ReadLine
        ts.ReadLine
        result = ts.Line
        ts.Close
        ''')
        interp = Interpreter()
        interp.interpret(program)
        assert interp._environment.get('result') == 3

class TestFSODeleteCopyMove:
    """Test FileSystemObject Delete, Copy, Move operations."""

    def test_deletefile(self, tmp_path):
        fpath = tmp_path / "test_delete.txt"
        fpath.write_text("delete me")
        program = parse(f'''
        Dim fso
        Set fso = CreateObject("Scripting.FileSystemObject")
        fso.DeleteFile "{fpath}"
        result = fso.FileExists("{fpath}")
        ''')
        interp = Interpreter()
        interp.interpret(program)
        assert interp._environment.get('result') is False

    def test_copyfile(self, tmp_path):
        src = tmp_path / "source.txt"
        src.write_text("copy me")
        dest = tmp_path / "dest.txt"
        program = parse(f'''
        Dim fso
        Set fso = CreateObject("Scripting.FileSystemObject")
        fso.CopyFile "{src}", "{dest}"
        result = fso.FileExists("{dest}")
        ''')
        interp = Interpreter()
        interp.interpret(program)
        assert interp._environment.get('result') is True
        assert dest.read_text() == "copy me"

    def test_movefile(self, tmp_path):
        src = tmp_path / "move_src.txt"
        src.write_text("move me")
        dest = tmp_path / "move_dest.txt"
        program = parse(f'''
        Dim fso
        Set fso = CreateObject("Scripting.FileSystemObject")
        fso.MoveFile "{src}", "{dest}"
        srcExists = fso.FileExists("{src}")
        destExists = fso.FileExists("{dest}")
        ''')
        interp = Interpreter()
        interp.interpret(program)
        assert interp._environment.get('srcExists') is False
        assert interp._environment.get('destExists') is True

    def test_createfolder(self, tmp_path):
        folder = tmp_path / "new_folder"
        program = parse(f'''
        Dim fso
        Set fso = CreateObject("Scripting.FileSystemObject")
        fso.CreateFolder "{folder}"
        result = fso.FolderExists("{folder}")
        ''')
        interp = Interpreter()
        interp.interpret(program)
        assert interp._environment.get('result') is True

    def test_deletefolder(self, tmp_path):
        folder = tmp_path / "del_folder"
        folder.mkdir()
        program = parse(f'''
        Dim fso
        Set fso = CreateObject("Scripting.FileSystemObject")
        fso.DeleteFolder "{folder}"
        result = fso.FolderExists("{folder}")
        ''')
        interp = Interpreter()
        interp.interpret(program)
        assert interp._environment.get('result') is False

    def test_copyfolder(self, tmp_path):
        src = tmp_path / "copy_src_dir"
        src.mkdir()
        (src / "file.txt").write_text("hello")
        dest = tmp_path / "copy_dest_dir"
        program = parse(f'''
        Dim fso
        Set fso = CreateObject("Scripting.FileSystemObject")
        fso.CopyFolder "{src}", "{dest}"
        result = fso.FolderExists("{dest}")
        ''')
        interp = Interpreter()
        interp.interpret(program)
        assert interp._environment.get('result') is True
        assert (dest / "file.txt").read_text() == "hello"

class TestFSOGetFileFolder:
    """Test FileSystemObject GetFile and GetFolder objects."""

    def test_getfile_properties(self, tmp_path):
        fpath = tmp_path / "props.txt"
        fpath.write_text("hello world")
        program = parse(f'''
        Dim fso, f
        Set fso = CreateObject("Scripting.FileSystemObject")
        Set f = fso.GetFile("{fpath}")
        name = f.Name
        size = f.Size
        ''')
        interp = Interpreter()
        interp.interpret(program)
        assert interp._environment.get('name') == 'props.txt'
        assert interp._environment.get('size') == 11

    def test_getfile_path(self, tmp_path):
        fpath = tmp_path / "pathtest.txt"
        fpath.write_text("")
        program = parse(f'''
        Dim fso, f
        Set fso = CreateObject("Scripting.FileSystemObject")
        Set f = fso.GetFile("{fpath}")
        result = f.Path
        ''')
        interp = Interpreter()
        interp.interpret(program)
        assert interp._environment.get('result') == str(fpath)

    def test_getfile_nonexistent_error(self):
        program = parse('''
        Dim fso, f
        Set fso = CreateObject("Scripting.FileSystemObject")
        Set f = fso.GetFile("/nonexistent_xyz.txt")
        ''')
        interp = Interpreter()
        with pytest.raises(VBScriptError, match="not found"):
            interp.interpret(program)

    def test_getfolder_properties(self, tmp_path):
        program = parse(f'''
        Dim fso, f
        Set fso = CreateObject("Scripting.FileSystemObject")
        Set f = fso.GetFolder("{tmp_path}")
        name = f.Name
        ftype = f.Type
        ''')
        interp = Interpreter()
        interp.interpret(program)
        assert interp._environment.get('name') == tmp_path.name
        assert interp._environment.get('ftype') == 'File Folder'

    def test_getfolder_nonexistent_error(self):
        program = parse('''
        Dim fso, f
        Set fso = CreateObject("Scripting.FileSystemObject")
        Set f = fso.GetFolder("/nonexistent_dir_xyz")
        ''')
        interp = Interpreter()
        with pytest.raises(VBScriptError, match="not found"):
            interp.interpret(program)

    def test_file_delete(self, tmp_path):
        fpath = tmp_path / "file_del.txt"
        fpath.write_text("delete")
        program = parse(f'''
        Dim fso, f
        Set fso = CreateObject("Scripting.FileSystemObject")
        Set f = fso.GetFile("{fpath}")
        f.Delete
        result = fso.FileExists("{fpath}")
        ''')
        interp = Interpreter()
        interp.interpret(program)
        assert interp._environment.get('result') is False

    def test_file_openastextstream(self, tmp_path):
        fpath = tmp_path / "open_as.txt"
        fpath.write_text("content")
        program = parse(f'''
        Dim fso, f, ts
        Set fso = CreateObject("Scripting.FileSystemObject")
        Set f = fso.GetFile("{fpath}")
        Set ts = f.OpenAsTextStream(1)
        result = ts.ReadAll
        ts.Close
        ''')
        interp = Interpreter()
        interp.interpret(program)
        assert interp._environment.get('result') == 'content'

    def test_file_parentfolder(self, tmp_path):
        fpath = tmp_path / "parent_test.txt"
        fpath.write_text("")
        program = parse(f'''
        Dim fso, f
        Set fso = CreateObject("Scripting.FileSystemObject")
        Set f = fso.GetFile("{fpath}")
        result = f.ParentFolder.Name
        ''')
        interp = Interpreter()
        interp.interpret(program)
        assert interp._environment.get('result') == tmp_path.name

class TestFSOFolderCollections:
    """Test Folder.Files and Folder.SubFolders collections."""

    def test_files_count(self, tmp_path):
        (tmp_path / "a.txt").write_text("a")
        (tmp_path / "b.txt").write_text("b")
        program = parse(f'''
        Dim fso, folder
        Set fso = CreateObject("Scripting.FileSystemObject")
        Set folder = fso.GetFolder("{tmp_path}")
        result = folder.Files.Count
        ''')
        interp = Interpreter()
        interp.interpret(program)
        assert interp._environment.get('result') == 2

    def test_subfolders_count(self, tmp_path):
        (tmp_path / "sub1").mkdir()
        (tmp_path / "sub2").mkdir()
        program = parse(f'''
        Dim fso, folder
        Set fso = CreateObject("Scripting.FileSystemObject")
        Set folder = fso.GetFolder("{tmp_path}")
        result = folder.SubFolders.Count
        ''')
        interp = Interpreter()
        interp.interpret(program)
        assert interp._environment.get('result') == 2

    def test_foreach_files(self, tmp_path):
        (tmp_path / "x.txt").write_text("x")
        (tmp_path / "y.txt").write_text("y")
        output = io.StringIO()
        program = parse(f'''
        Dim fso, folder, f, names
        Set fso = CreateObject("Scripting.FileSystemObject")
        Set folder = fso.GetFolder("{tmp_path}")
        names = ""
        For Each f In folder.Files
            If names <> "" Then names = names & ","
            names = names & f.Name
        Next
        WScript.Echo names
        ''')
        interp = Interpreter(output_stream=output)
        interp.interpret(program)
        result = output.getvalue().strip()
        parts = sorted(result.split(','))
        assert parts == ['x.txt', 'y.txt']

class TestFSOSpecialFolder:
    """Test GetSpecialFolder."""

    def test_temp_folder(self):
        program = parse('''
        Dim fso, folder
        Set fso = CreateObject("Scripting.FileSystemObject")
        Set folder = fso.GetSpecialFolder(2)
        result = fso.FolderExists(folder.Path)
        ''')
        interp = Interpreter()
        interp.interpret(program)
        assert interp._environment.get('result') is True

    def test_invalid_folder_error(self):
        program = parse('''
        Dim fso
        Set fso = CreateObject("Scripting.FileSystemObject")
        Set folder = fso.GetSpecialFolder(99)
        ''')
        interp = Interpreter()
        with pytest.raises(VBScriptError):
            interp.interpret(program)

class TestFSOIntegration:
    """Integration tests combining multiple FSO operations."""

    def test_create_write_read_delete(self, tmp_path):
        fpath = tmp_path / "integration.txt"
        output = io.StringIO()
        program = parse(f'''
        Dim fso, ts
        Set fso = CreateObject("Scripting.FileSystemObject")

        Set ts = fso.CreateTextFile("{fpath}")
        ts.WriteLine "Hello"
        ts.WriteLine "World"
        ts.Close

        Set ts = fso.OpenTextFile("{fpath}")
        Dim line1, line2
        line1 = ts.ReadLine
        line2 = ts.ReadLine
        ts.Close

        WScript.Echo line1
        WScript.Echo line2

        fso.DeleteFile "{fpath}"
        WScript.Echo fso.FileExists("{fpath}")
        ''')
        interp = Interpreter(output_stream=output)
        interp.interpret(program)
        lines = output.getvalue().strip().split('\n')
        assert lines == ['Hello', 'World', 'False']

    def test_folder_with_files(self, tmp_path):
        folder = tmp_path / "test_dir"
        program = parse(f'''
        Dim fso
        Set fso = CreateObject("Scripting.FileSystemObject")
        fso.CreateFolder "{folder}"

        Dim ts
        Set ts = fso.CreateTextFile(fso.BuildPath("{folder}", "file1.txt"))
        ts.Write "content1"
        ts.Close

        Set ts = fso.CreateTextFile(fso.BuildPath("{folder}", "file2.txt"))
        ts.Write "content2"
        ts.Close

        Dim f
        Set f = fso.GetFolder("{folder}")
        result = f.Files.Count
        ''')
        interp = Interpreter()
        interp.interpret(program)
        assert interp._environment.get('result') == 2

    def test_file_date_properties(self, tmp_path):
        fpath = tmp_path / "date_test.txt"
        fpath.write_text("test")
        program = parse(f'''
        Dim fso, f
        Set fso = CreateObject("Scripting.FileSystemObject")
        Set f = fso.GetFile("{fpath}")
        result = TypeName(f.DateCreated)
        ''')
        interp = Interpreter()
        interp.interpret(program)
        assert interp._environment.get('result') == 'Date'

    def test_drives_property(self):
        program = parse('''
        Dim fso
        Set fso = CreateObject("Scripting.FileSystemObject")
        result = (fso.Drives.Count >= 1)
        ''')
        interp = Interpreter()
        interp.interpret(program)
        assert interp._environment.get('result') is True
