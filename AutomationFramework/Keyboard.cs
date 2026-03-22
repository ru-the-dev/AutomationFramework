namespace AutomationFramework;

/// <summary>
/// Simulates keyboard input using Win32 SendInput.
/// </summary>
public sealed class Keyboard
{

    /// <summary>
    /// Configuration values that control keyboard input behavior.
    /// </summary>
    public sealed record class Options
    {
        public float MinDurationFactor { get; init; } = 0.9f;
		public float MaxDurationFactor { get; init; } = 1.2f;

        public TimeSpan DefaultKeyPressDuration { get; init; } = TimeSpan.FromMilliseconds(80);

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
            
            if (DefaultKeyPressDuration < TimeSpan.Zero)
			{
				throw new ArgumentOutOfRangeException(nameof(DefaultKeyPressDuration), "DefaultKeyPressDuration cannot be negative.");
			}
        }
    }

    private readonly Options _options;

    private readonly SemaphoreSlim _semaphore = new SemaphoreSlim(1, 1);

    public Keyboard(Options? options = null)
    {
        _options = options ?? new Options();
        _options.Validate();
    }

    /// <summary>
    /// Sends a key down event for a virtual-key code.
    /// </summary>
    public Task KeyDownAsync(VirtualKey virtualKey, CancellationToken cancellationToken = default)
    {
        return KeyDownAsync((ushort)virtualKey, cancellationToken);
    }

    /// <summary>
    /// Sends a key down event for a virtual-key code.
    /// </summary>
    public async Task KeyDownAsync(ushort virtualKey, CancellationToken cancellationToken = default)
    {
        await _semaphore.WaitAsync(cancellationToken).ConfigureAwait(false);
        try
        {
            InputNative.SendKeyDown(virtualKey);
        }
        finally
        {
            _semaphore.Release();
        }
    }

    /// <summary>
    /// Sends a key up event for a virtual-key code.
    /// </summary>
    public Task KeyUpAsync(VirtualKey virtualKey, CancellationToken cancellationToken = default)
    {
        return KeyUpAsync((ushort)virtualKey, cancellationToken);
    }

    /// <summary>
    /// Sends a key up event for a virtual-key code.
    /// </summary>
    public async Task KeyUpAsync(ushort virtualKey, CancellationToken cancellationToken = default)
    {
        await _semaphore.WaitAsync(cancellationToken).ConfigureAwait(false);
        try
        {
            InputNative.SendKeyUp(virtualKey);
        }
        finally
        {
            _semaphore.Release();
        }
    }

    /// <summary>
    /// Sends a full key press (down + up) for a virtual-key code.
    /// </summary>
    public Task PressKeyAsync(
        VirtualKey virtualKey,
        TimeSpan? holdDuration = null,
        CancellationToken cancellationToken = default)
    {
        return PressKeyAsync((ushort)virtualKey, holdDuration, cancellationToken);
    }

    /// <summary>
    /// Sends a key down event, waits, then sends key up.
    /// </summary>
    public async Task PressKeyAsync(
        ushort virtualKey,
        TimeSpan? holdDuration = null,
        CancellationToken cancellationToken = default)
    {
        await _semaphore.WaitAsync(cancellationToken).ConfigureAwait(false);
        try
        {
            holdDuration ??= _options.DefaultKeyPressDuration;
            holdDuration = TimespanExtensions.ApplyRandomFactor(holdDuration.Value, _options.MinDurationFactor, _options.MaxDurationFactor);

            InputNative.SendKeyDown(virtualKey);

            if (holdDuration > TimeSpan.Zero)
            {
                await Task.Delay(holdDuration.Value, cancellationToken).ConfigureAwait(false);
            }

            InputNative.SendKeyUp(virtualKey);
        }
        finally
        {
            _semaphore.Release();
        }
    }


    /// <summary>
    /// Presses a key chord and holds it before release.
    /// </summary>
    public async Task PressChordAsync(
        VirtualKey[] virtualKeys,
        TimeSpan? keypressInterval = null,
        TimeSpan? holdDuration = null,
        CancellationToken cancellationToken = default
    )
    {
        ArgumentNullException.ThrowIfNull(virtualKeys);
        if (virtualKeys.Length == 0)
        {
            throw new ArgumentException("At least one key is required.", nameof(virtualKeys));
        }

        await _semaphore.WaitAsync(cancellationToken).ConfigureAwait(false);
        try
        {

            holdDuration ??= _options.DefaultKeyPressDuration;
            keypressInterval ??= _options.DefaultKeyPressDuration;

            foreach (var key in virtualKeys)
            {
                InputNative.SendKeyDown((ushort)key);
                var interval = keypressInterval.Value.ApplyRandomFactor(_options.MinDurationFactor, _options.MaxDurationFactor);
                await Task.Delay(interval, cancellationToken).ConfigureAwait(false);
            }

            if (holdDuration > TimeSpan.Zero)
            {
                var durationRandomized = holdDuration.Value.ApplyRandomFactor(_options.MinDurationFactor, _options.MaxDurationFactor);
                await Task.Delay(durationRandomized, cancellationToken).ConfigureAwait(false);
            }

            for (var index = virtualKeys.Length - 1; index >= 0; index--)
            {
                // TODO: consider adding a small random delay between key releases to better simulate human behavior.
                InputNative.SendKeyUp((ushort)virtualKeys[index]);
                var interval = keypressInterval.Value.ApplyRandomFactor(_options.MinDurationFactor, _options.MaxDurationFactor) * 0.5f;
                await Task.Delay(interval, cancellationToken).ConfigureAwait(false);
            }
        }
        finally
        {
            _semaphore.Release();
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

        await _semaphore.WaitAsync(cancellationToken).ConfigureAwait(false);
        try
        {

            charDelay ??= _options.DefaultKeyPressDuration;

            foreach (var character in text)
            {
                cancellationToken.ThrowIfCancellationRequested();
                InputNative.SendUnicodeChar(character);

                if (charDelay > TimeSpan.Zero)
                {
                    var randomizedDelay = charDelay.Value.ApplyRandomFactor(_options.MinDurationFactor, _options.MaxDurationFactor);
                    await Task.Delay(randomizedDelay, cancellationToken).ConfigureAwait(false);
                }
            }
        }
        finally
        {
            _semaphore.Release();
        }
    }
}