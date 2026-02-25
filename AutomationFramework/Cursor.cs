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

		var controlPoints = CreateControlPoints(startPos, targetPos, distance, _options.CurvaturePixels);
		
		var totalTicks = adjustedDuration.Ticks;

		// Track consumed delay time so total movement stays near adjustedDuration.
		var assignedTicks = 0L;

		for (var index = 1; index <= steps; index++)
		{
			cancellationToken.ThrowIfCancellationRequested();

			// Sample the cubic Bezier curve at normalized progress t.
			var t = index / (float)steps;
			
			// point at T along bezier curve before jitter.
			var tPoint = CubicBezier(startPos, controlPoints.C1, controlPoints.C2, targetPos, t);

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

			SetCursorPos((int)Math.Round(tPoint.X), (int)Math.Round(tPoint.Y));

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

	//TODO: probably should be in a seperate class since these are general purpose mathematical extensions, not specific to cursor movement.
	/// <summary>
	/// Creates two randomized control points for a cubic Bezier path.
	/// </summary>
	private static (Vector2 C1, Vector2 C2) CreateControlPoints(
		Vector2 start,
		Vector2 end,
		double distance,
		double curvaturePixels)
	{
		// displacement from start to end.
		var displacement = end - start;

		if (distance <= double.Epsilon)
			return (start, end);

		var direction = Vector2.Normalize(displacement);

		var perpendicular = new Vector2(-direction.Y, direction.X);
		
		// Curvature magnitude: scales with distance, clamped by config.
		var curve = (float)Math.Min(curvaturePixels, Math.Max(8.0, distance * 0.35));
		var alongJitter = curve * 0.2f;

		// Base control points near 1/3 and 2/3 of the segment.
		var c1Base = start + displacement * 0.33f;
		var c2Base = start + displacement * 0.66f;

		// Add randomized sideways bend + slight forward/backward jitter.
		var c1 = c1Base
			+ perpendicular * Random.Shared.NextFloat(-curve, curve)
			+ direction * Random.Shared.NextFloat(-alongJitter, alongJitter);

		var c2 = c2Base
			+ perpendicular * Random.Shared.NextFloat(-curve, curve)
			+ direction * Random.Shared.NextFloat(-alongJitter, alongJitter);
			
		return (c1, c2);
	}


	//TODO: probably should be in a seperate class since these are general purpose mathematical extensions, not specific to cursor movement.
	/// <summary>
	/// Evaluates a cubic Bezier curve at t in [0, 1].
	/// </summary>
	private static Vector2 CubicBezier(
		Vector2 p0,
		Vector2 p1,
		Vector2 p2,
		Vector2 p3,
		float t)
	{;
		var oneMinusT = 1f - t;
		var oneMinusTSquared = oneMinusT * oneMinusT;
		var tSquared = t * t;

		return
			(oneMinusTSquared * oneMinusT * p0) +
			(3f * oneMinusTSquared * t * p1) +
			(3f * oneMinusT * tSquared * p2) +
			(tSquared * t * p3);
	}


	// Win32 API: move cursor to absolute screen coordinates.
	[DllImport("user32.dll")]
	private static extern bool SetCursorPos(int x, int y);

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
