# Third-party notices

## OpenRazer protocol references

MysticFight's minimal Razer Ornata V3 HID implementation was independently
written from the public protocol behavior documented by these projects:

- OpenRazer: https://github.com/openrazer/openrazer
- openrazer-win: https://github.com/Ar4ikov/openrazer-win

Both projects are licensed under GPL-2.0-or-later. MysticFight is distributed
under GPLv3, a compatible later version of the GNU General Public License.

MysticFight does not bundle or load either project at runtime. The implementation
is intentionally restricted to detection of Razer VID `0x1532`, Ornata V3 PIDs
`0x02A1` and `0x028F`, and the volatile extended custom-frame commands needed
for ten RGB zones. It contains no firmware, bootloader, EEPROM, macro, polling,
or persistent-lighting command path.
