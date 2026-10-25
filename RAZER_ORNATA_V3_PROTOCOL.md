# Razer Ornata V3 protocol notes

Research snapshot: 2026-10-25.

Primary references:

- OpenRazer: https://github.com/openrazer/openrazer
- openrazer-win: https://github.com/Ar4ikov/openrazer-win

## Verified device identity and interface

- USB vendor ID: `0x1532` (Razer).
- Product ID `0x02A1`: `RazerOrnataV3` in current OpenRazer.
- Product ID `0x028F`: `RazerOrnataV3_Alternate` in current OpenRazer.
- Both variants declare a `1 x 10` lighting matrix.
- Both recipes address keyboard interface index 2 with a nominal 600 us wait.
- On Windows, interface 2 is selected by `&mi_02` in the HID device path and
  must expose a 91-byte feature buffer: one report-ID byte followed by the
  90-byte Razer control report.

MysticFight requires all four checks (VID, whitelisted PID, `MI_02`, feature
length 91). Detection opens only a zero-access inspection handle and sends no
feature report.

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

## Volatile ten-zone frame used by MysticFight

Both Ornata V3 PIDs use transaction ID `0x1F` for these recipes.

1. Extended custom row: class `0x0F`, command `0x03`, data size `0x47`.
   Arguments 2, 3 and 4 are row `0`, first column `0` and last column `9`.
   Arguments 5 through 34 contain ten consecutive `R,G,B` triples.
2. Extended custom effect: class `0x0F`, command `0x02`, data size `0x0C`.
   Arguments 0 and 1 are zero (no variable storage); argument 2 is custom
   effect `0x08`.

The row is uploaded first and the custom effect is selected once after opening
the device. Later frames need only the row command. MysticFight exposes no
generic command API: its write whitelist accepts exactly these two report
shapes and validates their zero/reserved fields and checksum again immediately
before sending.

The OpenRazer static-effect recipe uses variable storage (`VARSTORE=1`) for
these devices. Because that may be persistent, MysticFight deliberately does
not implement it. A uniform visually static color is produced by repeating one
RGB triple across the volatile ten-zone custom frame.

## Timing and safety decisions

- OpenRazer specifies 600 us between request and response; openrazer-win uses a
  1 ms Windows minimum. MysticFight follows the 1 ms minimum.
- Razer output defaults to 15 FPS, independently selectable as 15, 20 or 25
  FPS. Transfer latency is measured; operations exceeding the frame budget are
  logged.
- The Razer worker is separate from the UI, sensor and MSI SDK workers. It owns
  its HID handle, samples only the newest temperature, queues no frames and
  suppresses identical frames.
- Before dynamic control can be enabled, Settings requires a separate manual
  one-color safety test. Its confirmation names the exact color; the worker
  sends one uniform frame, closes the HID handle, records success, and leaves
  every dynamic Razer profile and consent flag disabled.
- After that test, no dynamic Razer write is possible unless a profile enables
  Razer, the second explicit warning has been accepted and saved, the verified
  keyboard is present, and a valid sensor color (or an explicit global OFF
  request) exists.
- Disconnects close the handle and return to read-only detection. Application
  exit only closes the handle; it sends no shutdown color or persistent command.
- Windows `HidD_SetFeature`/`HidD_GetFeature` calls are synchronous. They are
  isolated from MSI and the UI, bounded by the protocol retry count, and the
  application retains its existing bounded worker-shutdown policy. No claim is
  made that Windows provides a cancellable per-call HID deadline here.

Not implemented: firmware, bootloader, EEPROM, macros, keyboard polling,
persistent lighting, driver replacement, services, or automatic hardware-write
tests.
