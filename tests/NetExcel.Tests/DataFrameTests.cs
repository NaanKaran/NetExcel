using FluentAssertions;

namespace NetXLCsvTests;

/// <summary>Unit tests for the DataFrame API (filter, select, sort, group, add/remove columns).</summary>
public sealed class DataFrameTests
{
    // ── Helpers ────────────────────────────────────────────────────────────────

    private static NetDataFrame MakePeopleFrame() =>
        DataFrame.FromList(new[]
        {
            new { Name = "Alice", Age = 30, Country = "India" },
            new { Name = "Bob",   Age = 17, Country = "UK"    },
            new { Name = "Carol", Age = 25, Country = "India" },
            new { Name = "Dave",  Age = 15, Country = "USA"   },
            new { Name = "Eve",   Age = 42, Country = "UK"    }
        });

    // ── FromList ───────────────────────────────────────────────────────────────

    [Fact]
    public void FromList_CreatesCorrectShape()
    {
        var df = MakePeopleFrame();
        df.RowCount.Should().Be(5);
        df.ColumnCount.Should().Be(3);
        df.Schema.Columns.Select(c => c.Name).Should().Equal("Name", "Age", "Country");
    }

    // ── Filter ─────────────────────────────────────────────────────────────────

    [Fact]
    public void Filter_ShouldReturnOnlyMatchingRows()
    {
        var df = DataFrame.FromList(new[]
        {
            new { Name = "John",  Age = 25 },
            new { Name = "Alice", Age = 15 }
        });

        var result = df.Filter(r => (int)r["Age"]! > 18);

        result.RowCount.Should().Be(1);
        result.GetValue(0, 0).Should().Be("John");
    }

    [Fact]
    public void Filter_NoMatchReturnsEmptyFrame()
    {
        var df = MakePeopleFrame();
        var result = df.Filter(r => (int)r["Age"]! > 100);
        result.RowCount.Should().Be(0);
    }

    [Fact]
    public void Filter_AllMatchReturnsFullFrame()
    {
        var df = MakePeopleFrame();
        var result = df.Filter(r => (int)r["Age"]! > 0);
        result.RowCount.Should().Be(5);
    }

    [Fact]
    public void Filter_IsImmutable_OriginalUnchanged()
    {
        var df = MakePeopleFrame();
        var filtered = df.Filter(r => (int)r["Age"]! > 18);

        df.RowCount.Should().Be(5);     // original untouched
        filtered.RowCount.Should().Be(3);
    }

    // ── Select ─────────────────────────────────────────────────────────────────

    [Fact]
    public void Select_ReturnsSubsetOfColumns()
    {
        var df = MakePeopleFrame();
        var result = df.Select("Name", "Country");

        result.ColumnCount.Should().Be(2);
        result.Schema.Columns.Select(c => c.Name).Should().Equal("Name", "Country");
    }

    [Fact]
    public void Select_UnknownColumnThrows()
    {
        var df = MakePeopleFrame();
        var act = () => df.Select("Name", "DoesNotExist");
        act.Should().Throw<KeyNotFoundException>();
    }

    // ── SortBy ─────────────────────────────────────────────────────────────────

    [Fact]
    public void SortBy_AscendingAge()
    {
        var df = MakePeopleFrame();
        var sorted = df.SortBy("Age");
        // IDataFrame : IEnumerable<DataRow> — iterate via LINQ directly on the sorted frame
        var ages = sorted.AsEnumerable().Select(r => Convert.ToInt32(r["Age"])).ToList();

        ages.Should().BeInAscendingOrder();
    }

    [Fact]
    public void SortBy_DescendingAge()
    {
        var df = MakePeopleFrame();
        var sorted = df.SortBy("Age", ascending: false);
        var firstAge = (int)sorted.GetValue(0, 1)!;
        firstAge.Should().Be(42); // Eve is oldest
    }

    // ── AddColumn ──────────────────────────────────────────────────────────────

    [Fact]
    public void AddColumn_AppendsNewColumn()
    {
        var df = MakePeopleFrame();
        var result = df.AddColumn("Active", true);

        result.ColumnCount.Should().Be(4);
        result.Schema.HasColumn("Active").Should().BeTrue();
        result.GetValue(0, 3).Should().Be(true);
    }

    [Fact]
    public void AddColumn_ReplacesExistingColumn()
    {
        var df = MakePeopleFrame();
        var result = df.AddColumn("Country", "Global");

        result.ColumnCount.Should().Be(3); // no extra column
        result.GetValue(0, 2).Should().Be("Global");
    }

    // ── RemoveColumn ───────────────────────────────────────────────────────────

    [Fact]
    public void RemoveColumn_DecrementsColumnCount()
    {
        var df = MakePeopleFrame();
        var result = df.RemoveColumn("Country");

        result.ColumnCount.Should().Be(2);
        result.Schema.HasColumn("Country").Should().BeFalse();
    }

    // ── GroupBy ────────────────────────────────────────────────────────────────

    [Fact]
    public void GroupBy_ReturnsTwoGroupsForCountry()
    {
        var df = MakePeopleFrame();
        var groups = df.GroupBy("Country");

        groups.Should().HaveCount(3); // India, UK, USA
        groups["India"].RowCount.Should().Be(2);
        groups["UK"].RowCount.Should().Be(2);
        groups["USA"].RowCount.Should().Be(1);
    }

    // ── Enumeration ────────────────────────────────────────────────────────────

    [Fact]
    public void Enumerate_YieldsAllRows()
    {
        var df = MakePeopleFrame();
        var count = df.Count();
        count.Should().Be(5);
    }

    // ── Empty ──────────────────────────────────────────────────────────────────

    [Fact]
    public void Empty_HasZeroRowsAndColumns()
    {
        var df = DataFrame.Empty();
        df.RowCount.Should().Be(0);
        df.ColumnCount.Should().Be(0);
    }

    // ── FromColumns ────────────────────────────────────────────────────────────

    [Fact]
    public void FromColumns_CorrectShape()
    {
        var df = DataFrame.FromColumns(new()
        {
            ["Id"]   = [1, 2, 3],
            ["Name"] = ["A", "B", "C"]
        });

        df.RowCount.Should().Be(3);
        df.ColumnCount.Should().Be(2);
    }

    [Fact]
    public void FromColumns_MismatchedLengthsThrows()
    {
        var act = () => DataFrame.FromColumns(new()
        {
            ["Id"]   = [1, 2],
            ["Name"] = ["A", "B", "C"]  // different length
        });
        act.Should().Throw<ArgumentException>();
    }
}
