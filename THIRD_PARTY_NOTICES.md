# Third-party notices

## OpenRazer protocol references

MysticFight's minimal Razer HID implementation was independently written from
the public protocol behavior documented by these projects:

- OpenRazer: https://github.com/openrazer/openrazer
- openrazer-win: https://github.com/Ar4ikov/openrazer-win

Both projects are licensed under GPL-2.0-or-later. MysticFight is distributed
under GPLv3, a compatible later version of the GNU General Public License.

MysticFight does not bundle or load either project at runtime. Its read-only HID
scan publishes only compatible devices, and its write path is intentionally
restricted to the exact whitelisted PIDs `0x0244`, `0x028F`, `0x0294`, `0x02A1`,
`0x02A2` and `0x02A3`. It implements only the volatile extended custom-frame
commands needed for each model's documented single-row lighting geometry and the
two documented persistent hand-off states (Off and Spectrum). It contains no
firmware, bootloader, EEPROM, macro, keyboard-polling, or arbitrary-command path.
