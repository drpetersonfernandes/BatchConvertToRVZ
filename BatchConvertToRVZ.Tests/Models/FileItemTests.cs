using BatchConvertToRVZ.Models;
using Xunit;

namespace BatchConvertToRVZ.Tests.Models;

public class FileItemTests
{
    [Fact]
    public void IsSelectedSetValueRaisesPropertyChanged()
    {
        var item = new FileItem();
        var raisedProperties = new List<string>();
        item.PropertyChanged += (_, e) =>
        {
            if (e.PropertyName != null) raisedProperties.Add(e.PropertyName);
        };

        item.IsSelected = false;

        Assert.Single(raisedProperties);
        Assert.Equal(nameof(FileItem.IsSelected), raisedProperties[0]);
        Assert.False(item.IsSelected);
    }

    [Fact]
    public void IsSelectedSetSameValueDoesNotRaisePropertyChanged()
    {
        var item = new FileItem { IsSelected = true };
        var raisedProperties = new List<string>();
        item.PropertyChanged += (_, e) =>
        {
            if (e.PropertyName != null) raisedProperties.Add(e.PropertyName);
        };

        item.IsSelected = true;

        Assert.Empty(raisedProperties);
    }

    [Fact]
    public void FileNameSetValueRaisesPropertyChanged()
    {
        var item = new FileItem();
        var raisedProperties = new List<string>();
        item.PropertyChanged += (_, e) =>
        {
            if (e.PropertyName != null) raisedProperties.Add(e.PropertyName);
        };

        item.FileName = "game.iso";

        Assert.Single(raisedProperties);
        Assert.Equal(nameof(FileItem.FileName), raisedProperties[0]);
        Assert.Equal("game.iso", item.FileName);
    }

    [Fact]
    public void FullPathSetValueRaisesPropertyChanged()
    {
        var item = new FileItem();
        var raisedProperties = new List<string>();
        item.PropertyChanged += (_, e) =>
        {
            if (e.PropertyName != null) raisedProperties.Add(e.PropertyName);
        };

        item.FullPath = @"C:\games\game.iso";

        Assert.Single(raisedProperties);
        Assert.Equal(nameof(FileItem.FullPath), raisedProperties[0]);
    }

    [Theory]
    [InlineData(512L, "512 B")]
    [InlineData(1024L, "1 KB")]
    [InlineData(1536L, "1.5 KB")]
    [InlineData(1048576L, "1 MB")]
    [InlineData(1073741824L, "1 GB")]
    [InlineData(1099511627776L, "1 TB")]
    public void FileSizeSetValueUpdatesDisplaySize(long bytes, string expectedDisplay)
    {
        var item = new FileItem();
        var raisedProperties = new List<string>();
        item.PropertyChanged += (_, e) =>
        {
            if (e.PropertyName != null) raisedProperties.Add(e.PropertyName);
        };

        item.FileSize = bytes;

        Assert.Equal(expectedDisplay, item.DisplaySize);
        Assert.Contains(nameof(FileItem.FileSize), raisedProperties);
        Assert.Contains(nameof(FileItem.DisplaySize), raisedProperties);
    }

    [Fact]
    public void FileItemDefaultValuesAreCorrect()
    {
        var item = new FileItem();

        Assert.True(item.IsSelected);
        Assert.Equal(string.Empty, item.FileName);
        Assert.Equal(string.Empty, item.FullPath);
        Assert.Equal(0L, item.FileSize);
        Assert.Equal("0 B", item.DisplaySize);
    }
}
