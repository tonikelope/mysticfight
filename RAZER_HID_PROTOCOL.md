<div align="justify">

# Razer HID protocol scope

Research snapshot: 2026-10-25.

Primary references:

- OpenRazer: https://github.com/openrazer/openrazer
- openrazer-win: https://github.com/Ar4ikov/openrazer-win (device/recipe snapshot
  `6a626b2d11069ec8be6a7eda6e0dd8ee3ece0998`)

## Explicit compatibility table

MysticFight does not assume that all Razer products use the same HID protocol.
Its write path accepts only the devices below, whose four required OpenRazer
recipes, transaction ID, feature-report size and single-row geometry were
cross-checked together.

This is the current implementation scope, not a claim that other Razer devices
cannot be supported. A new model can be added by defining and validating an
explicit per-PID profile for its commands, transaction ID, HID interface,
feature-report length and lighting geometry. It remains excluded from Settings
and from every write path until that validation is complete.

| USB PID | OpenRazer name | Matrix | Preferred HID interface |
| --- | --- | ---: | ---: |
| `0x0244` | Razer Tartarus Pro | `1 x 21` | 2 |
| `0x028F` | Razer Ornata V3 (Alternate) | `1 x 10` | 2 |
| `0x0294` | Razer Ornata V3 X | `1 x 1` | 1 |
| `0x02A1` | Razer Ornata V3 | `1 x 10` | 2 |
| `0x02A2` | Razer Ornata V3 X (Alternate) | `1 x 1` | 1 |
| `0x02A3` | Razer Ornata V3 Tenkeyless | `1 x 8` | 2 |

All six PIDs use a nominal 600 us wait, a 90-byte Razer report inside a
91-byte Windows feature report, transaction ID `0x1F`, and the same four
OpenRazer command recipes used by MysticFight. The geometry changes only the
last column and RGB payload length of the custom row.

Detection opens only zero-access inspection handles and sends no feature report.
Compatible physical devices are deduplicated by their Windows container ID plus
PID and shown by product name and PID. Models outside the table and supported
PIDs without the validated 91-byte control collection are omitted from Settings
and never receive commands. MysticFight does not infer compatibility from a
product name, vendor ID, report length, or partial recipe match.

The upstream interface index is a preference rather than the sole identity
check because Windows can expose the 91-byte control collection on another
`MI_XX` path. VID, explicitly whitelisted PID and feature length 91 are checked
again when the device is opened.

## 90-byte control report

The report following the Windows report-ID byte is:

| Offset | Size | Meaning |
| --- | ---: | --- |
| 0 | 1 | status |
| 1 | 1 | transaction ID |
| 2 | 2 | remaining packets, big-endian |
| 4 | 1 | protocol type |
| 5 | 1 | data size |
| 6 | 1 | command class |
| 7 | 1 | command ID |
| 8 | 80 | arguments |
| 88 | 1 | checksum |
| 89 | 1 | reserved |

The checksum is XOR over report bytes 2 through 87 inclusive. Responses are
accepted only if their checksum, remaining-packet fields, command class and
command ID match. Status `0x02` (successful) and `0x01` (busy/accepted) are
success; `0x05` (not supported) stops retries. Other failures use at most five
attempts with 10 ms between retries.

## Commands used by MysticFight

Every permitted report is constructed internally and validated against an exact
allowlist immediately before sending. There is no generic Razer command API.

1. Extended custom row: class `0x0F`, command `0x03`, data size `0x47`.
   Arguments 2, 3 and 4 are row `0`, first column `0` and the PID-specific last
   column (`0`, `7`, `9` or `20`). The following bytes are the corresponding
   `R,G,B` triples.
2. Volatile extended custom effect: class `0x0F`, command `0x02`, data size
   `0x0C`. Arguments 0 and 1 are zero; argument 2 is custom effect `0x08`.
3. Persistent Off/Spectrum hand-off: class `0x0F`, command `0x02`, data size
   `0x06`, with `VARSTORE=1` and `BACKLIGHT_LED=5`. Effect `0x00` is Off and
   effect `0x03` is Spectrum Cycling.

The row is uploaded first and the custom effect is selected once after opening
the device. Later colour changes need only one row command. Static colours are
not stored persistently: while MysticFight owns the device it uses the volatile
custom effect, and persistent commands are reserved for the configured Off or
Spectrum Cycling hand-off.

Profiles may share the backward-compatible General data source, sensor and
threshold/colour scheme, or assign an independent scheme to each selected
physical Razer device. LibreHardwareMonitor is downloaded and parsed only once
per polling cycle; every runtime reads its assigned sensor from that coherent
snapshot and keeps its own interpolation and FPS deadlines. Once its visible
colour settles, that device receives no repeated frame writes.

A release command is sent only to a compatible device that MysticFight
successfully controlled, when a new active profile no longer selects it or the
application exits. A committed Windows shutdown/restart uses the same bounded
release path before the HID handles are closed. Mere detection never sends a
release command.

## Timing and safety decisions

- The 600 us upstream delay is rounded to a 1 ms Windows minimum.
- Razer output defaults to 25 FPS and is independently selectable as 15, 20 or
  25 FPS per physical device. FPS is a maximum transition cadence, not a
  continuous write rate: identical frames are suppressed at both the worker and
  controller boundaries. LibreHardwareMonitor sampling continues independently;
  a settled Razer worker wakes only for a sensor/source change, configuration,
  device detection/retry or shutdown work.
- The Razer worker is separate from the UI, sensor and MSI SDK workers. It owns
  one controller per selected device, samples only the newest temperature and
  queues no stale frames.
- Every profile starts with no Razer device selected. Checking a compatible
  device and saving the profile is the complete opt-in.
- No Razer write is possible unless that exact physical device is selected in
  the active profile, its PID and 91-byte control collection are whitelisted,
  and a valid sensor colour or explicit global Off request exists.
- Disconnects close only that device's handle and return it to read-only
  detection. Definitive Windows removal/invalid-handle errors abort the current
  transfer immediately; other possibly transient errors keep the bounded retry
  policy. Reopen attempts and read-only detection are throttled to a five-second
  interval without a polling loop, including when a KVM leaves a stale device
  interface visible.
- Windows `HidD_SetFeature`/`HidD_GetFeature` calls are synchronous and isolated
  from MSI and the UI. No claim is made that Windows provides a cancellable
  per-call HID deadline.

Not implemented: firmware, bootloader, EEPROM, macros, keyboard polling,
arbitrary persistent lighting, driver replacement, services, or automatic
hardware-write tests.

</div>
