using System.Numerics;
using System.Runtime.InteropServices;
using AutomationFramework.Extentions;

namespace AutomationFramework;

/// <summary>
/// Simulates smooth, human-like cursor movement.
/// </summary>
public sealed class Cursor
{
    public record class Options
	{
		public float MinDurationFactor { get; init; } = 0.9f;
		public float MaxDurationFactor { get; init; } = 1.2f;
		public float CurvaturePixels { get; init; } = 3500;
		public float PathJitterPixels { get; init; } = 1.4f;
		public int MinSteps { get; init; } = 24;
		public int MaxSteps { get; init; } = 150;

        internal void Validate()
        {
            if (MinDurationFactor <= 0 || MaxDurationFactor <= 0)
            {
                throw new ArgumentOutOfRangeException("Duration factors must be greater than zero.");
            }

            if (MinDurationFactor > MaxDurationFactor)
            {
                throw new ArgumentException("MinDurationFactor cannot be greater than MaxDurationFactor.");
            }
        }

	}

	private readonly Options _options;

    

	/// <summary>
	/// Creates a mouse mover with default options that will be reused for each move call.
	/// </summary>
	public Cursor(Options? options = null)
	{
		_options = options ?? new Options();
		_options.Validate();
	}

	/// <summary>
	/// Gets the current cursor position in screen coordinates.
	/// </summary>
	public Vector2 GetCurrentPosition()
	{
		if (!GetCursorPos(out var point))
			throw new InvalidOperationException("Failed to read the current cursor position.");

		return new Vector2(point.X, point.Y);
	}

	/// <summary>
	/// Moves the cursor using this instance's configured options.
	/// </summary>
	public async Task MoveToAsync(
		Vector2 targetPos,
		TimeSpan duration,
		CancellationToken cancellationToken)
	{
		// Basic input validation.
		if (duration <= TimeSpan.Zero)
			throw new ArgumentOutOfRangeException(nameof(duration), "Duration must be greater than zero.");

		// Capture the current cursor position as the movement start point.
		if (!GetCursorPos(out var startPoint))
			throw new InvalidOperationException("Failed to read the current cursor position.");

		var adjustedDuration = Duration.ApplyRandomFactor(duration, _options.MinDurationFactor, _options.MaxDurationFactor);
		var startPos = new Vector2(startPoint.X, startPoint.Y);

		var distance = Vector2.Distance(startPos, targetPos);

		// If the target is effectively the same as the start, skip movement but still respect the duration.
		if (distance < float.Epsilon)
			return;

		
		var steps = (int)Math.Clamp(adjustedDuration.TotalMilliseconds / 16, _options.MinSteps, _options.MaxSteps);

		var controlPoints = CursorPathMath.CreateControlPoints(startPos, targetPos, distance, _options.CurvaturePixels);
		
		var totalTicks = adjustedDuration.Ticks;

		// Track consumed delay time so total movement stays near adjustedDuration.
		var assignedTicks = 0L;

		for (var index = 1; index <= steps; index++)
		{
			cancellationToken.ThrowIfCancellationRequested();

			// Sample the cubic Bezier curve at normalized progress t.
			var t = index / (float)steps;
			
			// point at T along bezier curve before jitter.
			var tPoint = CursorPathMath.CubicBezier(startPos, controlPoints.C1, controlPoints.C2, targetPos, t);

			// Add small, decaying randomness so motion looks less robotic.
			var jitterScale = Math.Max(0.0f, 1.0f - t);
			var jitter = _options.PathJitterPixels * jitterScale;
			tPoint.X += Random.Shared.NextFloat(-jitter, jitter);
			tPoint.Y += Random.Shared.NextFloat(-jitter, jitter);

			if (index == steps)
			{
				// Ensure exact final landing.
				tPoint = targetPos;
			}

			InputNative.SendMouseAbsolute((int)Math.Round(tPoint.X), (int)Math.Round(tPoint.Y));

			var remainingTicks = totalTicks - assignedTicks;
			
			if (index < steps && remainingTicks > 0)
			{
				// Split remaining time across remaining steps with slight timing jitter.
				var baselineTicks = remainingTicks / (steps - index + 1);
				var jitterFactor = Random.Shared.NextFloat(0.7f, 1.3f);
				var delayTicks = (long)Math.Clamp(baselineTicks * jitterFactor, TimeSpan.TicksPerMillisecond, remainingTicks);


				assignedTicks += delayTicks;
				await Task.Delay(TimeSpan.FromTicks(delayTicks), cancellationToken).ConfigureAwait(false);
			}
		}
	}

	/// <summary>
	/// Presses a mouse button down.
	/// </summary>
	public void MouseDown(MouseButton button = MouseButton.Left)
	{
		InputNative.SendMouseButtonDown(button);
	}

	/// <summary>
	/// Releases a mouse button.
	/// </summary>
	public void MouseUp(MouseButton button = MouseButton.Left)
	{
		InputNative.SendMouseButtonUp(button);
	}

	/// <summary>
	/// Performs a full click (down + up) for a mouse button.
	/// </summary>
	public void Click(MouseButton button = MouseButton.Left)
	{
		MouseDown(button);
		MouseUp(button);
	}

	/// <summary>
	/// Performs a click and optionally holds the button before release.
	/// </summary>
	public async Task ClickAsync(
		MouseButton button = MouseButton.Left,
		TimeSpan? holdDuration = null,
		CancellationToken cancellationToken = default)
	{
		MouseDown(button);

		if (holdDuration is { } hold && hold > TimeSpan.Zero)
		{
			await Task.Delay(hold, cancellationToken).ConfigureAwait(false);
		}

		MouseUp(button);
	}

	// Win32 API: read current cursor coordinates.
	[DllImport("user32.dll")]
	private static extern bool GetCursorPos(out Point point);

	/// <summary>
	/// Native POINT struct layout for GetCursorPos interop.
	/// </summary>
	[StructLayout(LayoutKind.Sequential)]
	private struct Point
	{
		public int X;
		public int Y;
	}
}

public enum MouseButton
{
	Left,
	Right,
	Middle
}
