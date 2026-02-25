namespace AutomationFramework;

/// <summary>
/// Simulates keyboard input using Win32 SendInput.
/// </summary>
public sealed class Keyboard
{
    public sealed record class Options
    {
        
    }

    /// <summary>
    /// Sends a key down event for a virtual-key code.
    /// </summary>
    public void KeyDown(VirtualKey virtualKey)
    {
        KeyDown((ushort)virtualKey);
    }

    /// <summary>
    /// Sends a key down event for a virtual-key code.
    /// </summary>
    public void KeyDown(ushort virtualKey)
    {
        InputNative.SendKeyDown(virtualKey);
    }

    /// <summary>
    /// Sends a key up event for a virtual-key code.
    /// </summary>
    public void KeyUp(VirtualKey virtualKey)
    {
        KeyUp((ushort)virtualKey);
    }

    /// <summary>
    /// Sends a key up event for a virtual-key code.
    /// </summary>
    public void KeyUp(ushort virtualKey)
    {
        InputNative.SendKeyUp(virtualKey);
    }

    /// <summary>
    /// Sends a full key press (down + up) for a virtual-key code.
    /// </summary>
    public void PressKey(VirtualKey virtualKey)
    {
        PressKey((ushort)virtualKey);
    }

    /// <summary>
    /// Sends a full key press (down + up) for a virtual-key code.
    /// </summary>
    public void PressKey(ushort virtualKey)
    {
        KeyDown(virtualKey);
        KeyUp(virtualKey);
    }

    /// <summary>
    /// Sends a full key press (down + up) for a virtual-key code.
    /// </summary>
    public Task PressKeyAsync(
        VirtualKey virtualKey,
        TimeSpan holdDuration,
        CancellationToken cancellationToken = default)
    {
        return PressKeyAsync((ushort)virtualKey, holdDuration, cancellationToken);
    }

    /// <summary>
    /// Sends a key down event, waits, then sends key up.
    /// </summary>
    public async Task PressKeyAsync(
        ushort virtualKey,
        TimeSpan holdDuration,
        CancellationToken cancellationToken = default)
    {
        if (holdDuration < TimeSpan.Zero)
        {
            throw new ArgumentOutOfRangeException(nameof(holdDuration), "Hold duration cannot be negative.");
        }

        KeyDown(virtualKey);
        if (holdDuration > TimeSpan.Zero)
        {
            await Task.Delay(holdDuration, cancellationToken).ConfigureAwait(false);
        }

        KeyUp(virtualKey);
    }

    /// <summary>
    /// Presses a key chord such as Control + C.
    /// </summary>
    public void PressChord(params VirtualKey[] virtualKeys)
    {
        ArgumentNullException.ThrowIfNull(virtualKeys);
        if (virtualKeys.Length == 0)
        {
            throw new ArgumentException("At least one key is required.", nameof(virtualKeys));
        }

        foreach (var key in virtualKeys)
        {
            KeyDown(key);
        }

        for (var index = virtualKeys.Length - 1; index >= 0; index--)
        {
            KeyUp(virtualKeys[index]);
        }
    }

    /// <summary>
    /// Presses a key chord and holds it before release.
    /// </summary>
    public async Task PressChordAsync(
        TimeSpan holdDuration,
        CancellationToken cancellationToken = default,
        params VirtualKey[] virtualKeys)
    {
        ArgumentNullException.ThrowIfNull(virtualKeys);
        if (virtualKeys.Length == 0)
        {
            throw new ArgumentException("At least one key is required.", nameof(virtualKeys));
        }

        if (holdDuration < TimeSpan.Zero)
        {
            throw new ArgumentOutOfRangeException(nameof(holdDuration), "Hold duration cannot be negative.");
        }

        foreach (var key in virtualKeys)
        {
            KeyDown(key);
        }

        if (holdDuration > TimeSpan.Zero)
        {
            await Task.Delay(holdDuration, cancellationToken).ConfigureAwait(false);
        }

        for (var index = virtualKeys.Length - 1; index >= 0; index--)
        {
            KeyUp(virtualKeys[index]);
        }
    }

    /// <summary>
    /// Types text as Unicode keyboard input.
    /// </summary>
    public void TypeText(string text)
    {
        ArgumentNullException.ThrowIfNull(text);

        foreach (var character in text)
        {
            InputNative.SendUnicodeChar(character);
        }
    }

    /// <summary>
    /// Types text as Unicode keyboard input with an optional per-character delay.
    /// </summary>
    public async Task TypeTextAsync(
        string text,
        TimeSpan? charDelay = null,
        CancellationToken cancellationToken = default)
    {
        ArgumentNullException.ThrowIfNull(text);

        var delay = charDelay ?? TimeSpan.Zero;
        if (delay < TimeSpan.Zero)
        {
            throw new ArgumentOutOfRangeException(nameof(charDelay), "Character delay cannot be negative.");
        }

        foreach (var character in text)
        {
            cancellationToken.ThrowIfCancellationRequested();
            InputNative.SendUnicodeChar(character);

            if (delay > TimeSpan.Zero)
            {
                await Task.Delay(delay, cancellationToken).ConfigureAwait(false);
            }
        }
    }
}