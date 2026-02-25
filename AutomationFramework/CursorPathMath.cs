using System.Numerics;
using AutomationFramework.Extentions;

namespace AutomationFramework;

internal static class CursorPathMath
{
	/// <summary>
	/// Creates two randomized control points for a cubic Bezier path.
	/// </summary>
	internal static (Vector2 C1, Vector2 C2) CreateControlPoints(
		Vector2 start,
		Vector2 end,
		double distance,
		double curvaturePixels)
	{
		var displacement = end - start;

		if (distance <= double.Epsilon)
			return (start, end);

		var direction = Vector2.Normalize(displacement);
		var perpendicular = new Vector2(-direction.Y, direction.X);

		var curve = (float)Math.Min(curvaturePixels, Math.Max(8.0, distance * 0.35));
		var alongJitter = curve * 0.2f;

		var c1Base = start + displacement * 0.33f;
		var c2Base = start + displacement * 0.66f;

		var c1 = c1Base
			+ perpendicular * Random.Shared.NextFloat(-curve, curve)
			+ direction * Random.Shared.NextFloat(-alongJitter, alongJitter);

		var c2 = c2Base
			+ perpendicular * Random.Shared.NextFloat(-curve, curve)
			+ direction * Random.Shared.NextFloat(-alongJitter, alongJitter);

		return (c1, c2);
	}

	/// <summary>
	/// Evaluates a cubic Bezier curve at t in [0, 1].
	/// </summary>
	internal static Vector2 CubicBezier(
		Vector2 p0,
		Vector2 p1,
		Vector2 p2,
		Vector2 p3,
		float t)
	{
		var oneMinusT = 1f - t;
		var oneMinusTSquared = oneMinusT * oneMinusT;
		var tSquared = t * t;

		return
			(oneMinusTSquared * oneMinusT * p0) +
			(3f * oneMinusTSquared * t * p1) +
			(3f * oneMinusT * tSquared * p2) +
			(tSquared * t * p3);
	}
}