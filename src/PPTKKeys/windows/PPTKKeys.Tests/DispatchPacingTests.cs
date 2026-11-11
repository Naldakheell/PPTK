using PPTKKeys;
using Xunit;

namespace PPTKKeys.Tests;

public class DispatchPacingTests
{
    [Theory]
    [InlineData("CTRL 2", true)]
    [InlineData("CTRL 3", true)]
    [InlineData("ALT 2", true)]
    [InlineData("ALT 3", true)]
    [InlineData("ALT .", false)]
    [InlineData("H +", false)]
    [InlineData("T", false)]
    [InlineData("CMD ADD COLUMN EXPAND", false)]
    [InlineData(null, false)]
    public void Only_the_keyboard_grid_commands_are_paced(string? token, bool grid) =>
        Assert.Equal(grid, DispatchPacing.IsGridToken(token));

    [Fact]
    public void A_command_right_after_a_grid_command_waits_out_the_gap()
    {
        Assert.Equal(100, DispatchPacing.WaitMs(lastWasGrid: true, elapsedSinceLastEnd: 0));
        Assert.Equal(70, DispatchPacing.WaitMs(true, 30));
        Assert.Equal(0, DispatchPacing.WaitMs(true, 100));
        Assert.Equal(0, DispatchPacing.WaitMs(true, 5000));
    }

    [Fact]
    public void Nothing_waits_after_other_commands()
    {
        Assert.Equal(0, DispatchPacing.WaitMs(lastWasGrid: false, elapsedSinceLastEnd: 0));
    }

    [Fact]
    public void Hold_to_repeat_spacing_is_not_slowed()
    {
        // Spacing never follows a grid command in a hold, so its own 50 ms floor still governs.
        Assert.Equal(0, DispatchPacing.WaitMs(DispatchPacing.IsGridToken("H +"), 0));
    }
}
