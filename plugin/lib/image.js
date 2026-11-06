// Image size from the file header, for placing pictures at their natural
// aspect ratio. PNG, JPEG and GIF, the formats Office embeds everywhere.

export function imageInfo (buf) {
  if (buf.length > 24 && buf.readUInt32BE(0) === 0x89504e47) {
    return { ext: 'png', width: buf.readUInt32BE(16), height: buf.readUInt32BE(20) }
  }
  if (buf.length > 10 && buf.toString('latin1', 0, 3) === 'GIF') {
    return { ext: 'gif', width: buf.readUInt16LE(6), height: buf.readUInt16LE(8) }
  }
  if (buf.length > 4 && buf[0] === 0xff && buf[1] === 0xd8) {
    let p = 2
    while (p + 9 < buf.length) {
      if (buf[p] !== 0xff) { p++; continue }
      const marker = buf[p + 1]
      if (marker === 0xd8 || marker === 0x01 || (marker >= 0xd0 && marker <= 0xd7)) { p += 2; continue }
      const len = buf.readUInt16BE(p + 2)
      // SOF0–SOF15 except DHT (C4), JPG (C8) and DAC (CC).
      if (marker >= 0xc0 && marker <= 0xcf && marker !== 0xc4 && marker !== 0xc8 && marker !== 0xcc) {
        return { ext: 'jpeg', width: buf.readUInt16BE(p + 7), height: buf.readUInt16BE(p + 5) }
      }
      p += 2 + len
    }
    return { ext: 'jpeg', width: 800, height: 600 }
  }
  return undefined
}

export const EMU_PER_PX = 9525 // at 96 dpi
export const EMU_PER_CM = 360000
