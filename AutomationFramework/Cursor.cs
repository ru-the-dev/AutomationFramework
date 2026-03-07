using System.Numerics;
using System.Runtime.InteropServices;
using AutomationFramework.Extensions;

namespace AutomationFramework;

/// <summary>
/// Simulates smooth, human-like cursor movement.
/// </summary>
public sealed class Cursor
{
    /// <summary>
    /// Configuration values that control cursor movement behavior.
    /// </summary>
    public record class Options
	{
		public float MinSpeedFactor { get; init; } = 0.8f;
		public float MaxSpeedFactor { get; init; } = 1.2f;
		public float MovePathCurvaturePixels { get; init; } = 3500;
		public float MovePathJitterPixels { get; init; } = 1.4f;
		public int MinMoveSteps { get; init; } = 24;
		public int MaxMoveSteps { get; init; } = 150;

		public TimeSpan DefaultClickHoldDuration { get; init; } = TimeSpan.FromMilliseconds(100);

        internal void Validate()
        {
            if (MinSpeedFactor <= 0 || MaxSpeedFactor <= 0)
            {
                throw new ArgumentOutOfRangeException("Speed factors must be greater than zero.");
            }

            if (MinSpeedFactor > MaxSpeedFactor)
            {
                throw new ArgumentException("MinSpeedFactor cannot be greater than MaxSpeedFactor.");
            }

			if (MovePathCurvaturePixels < 0)
			{
				throw new ArgumentOutOfRangeException(nameof(MovePathCurvaturePixels), "MovePathCurvaturePixels cannot be negative.");
			}

			if (MovePathJitterPixels < 0)
			{
				throw new ArgumentOutOfRangeException(nameof(MovePathJitterPixels), "MovePathJitterPixels cannot be negative.");
			}

			if (MinMoveSteps < 1)
			{
				throw new ArgumentOutOfRangeException(nameof(MinMoveSteps), "MinMoveSteps must be at least 1.");
			}

			if (MaxMoveSteps < MinMoveSteps)
			{
				throw new ArgumentException("MaxMoveSteps cannot be less than MinMoveSteps.");
			}

			if (DefaultClickHoldDuration < TimeSpan.Zero)
			{
				throw new ArgumentOutOfRangeException(nameof(DefaultClickHoldDuration), "DefaultClickHoldDuration cannot be negative.");
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
		float speedPixelsPerSecond = 1500f,
		CancellationToken cancellationToken = default)
	{
		if (speedPixelsPerSecond <= 0 || float.IsNaN(speedPixelsPerSecond) || float.IsInfinity(speedPixelsPerSecond))
			throw new ArgumentOutOfRangeException(nameof(speedPixelsPerSecond), "Speed must be a finite value greater than zero.");

		var startPos = GetCurrentPosition();
		var distance = Vector2.Distance(startPos, targetPos);

		if (distance < float.Epsilon)
			return;

		var durationTicks = Math.Max(
			1L,
			(long)Math.Round(TimeSpan.TicksPerSecond * (distance / speedPixelsPerSecond)));

		await MoveToAsync(targetPos, TimeSpan.FromTicks(durationTicks), cancellationToken).ConfigureAwait(false);
	}

	/// <summary>
	/// Moves the cursor using this instance's configured options.
	/// </summary>
	public async Task MoveToAsync(
		Vector2 targetPos,
		TimeSpan duration,
		CancellationToken cancellationToken = default)
	{
		// Basic input validation.
		if (duration <= TimeSpan.Zero)
			throw new ArgumentOutOfRangeException(nameof(duration), "Duration must be greater than zero.");

		// Capture the current cursor position as the movement start point.
		if (!GetCursorPos(out var startPoint))
			throw new InvalidOperationException("Failed to read the current cursor position.");

		var adjustedDuration = Duration.ApplyRandomFactor(duration, _options.MinSpeedFactor, _options.MaxSpeedFactor);
		var startPos = new Vector2(startPoint.X, startPoint.Y);

		var distance = Vector2.Distance(startPos, targetPos);

		// If the target is effectively the same as the start, skip movement but still respect the duration.
		if (distance < float.Epsilon)
			return;

		
		var steps = (int)Math.Clamp(adjustedDuration.TotalMilliseconds / 16, _options.MinMoveSteps, _options.MaxMoveSteps);

		var controlPoints = CursorPathMath.CreateControlPoints(startPos, targetPos, distance, _options.MovePathCurvaturePixels);
		
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
			var jitter = _options.MovePathJitterPixels * jitterScale;
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
				var minDelayTicks = Math.Min(TimeSpan.TicksPerMillisecond, remainingTicks);
				var delayTicks = (long)Math.Clamp(baselineTicks * jitterFactor, minDelayTicks, remainingTicks);


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
	/// Performs a click and optionally holds the button before release.
	/// </summary>
	public async Task ClickAsync(
		MouseButton button = MouseButton.Left,
		TimeSpan? holdDuration = null,
		CancellationToken cancellationToken = default)
	{
		// apply default if not specified
		holdDuration ??= _options.DefaultClickHoldDuration;

		// then apply random factor to whatever the hold duration is.
		holdDuration = Duration.ApplyRandomFactor(holdDuration.Value, _options.MinSpeedFactor, _options.MaxSpeedFactor);

		MouseDown(button);

		if (holdDuration > TimeSpan.Zero)
		{
			await Task.Delay(holdDuration.Value, cancellationToken).ConfigureAwait(false);
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
