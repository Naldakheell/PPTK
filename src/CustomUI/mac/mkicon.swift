// Draws the PPTK "Replace Fonts" ribbon icon: charcoal A, cyan replace arrow. Usage: mkicon <size> <out.png>
import AppKit
let size = CGFloat(Double(CommandLine.arguments[1])!)
let out = CommandLine.arguments[2]
let s = size / 16
let rep = NSBitmapImageRep(bitmapDataPlanes: nil, pixelsWide: Int(size), pixelsHigh: Int(size), bitsPerSample: 8,
                           samplesPerPixel: 4, hasAlpha: true, isPlanar: false, colorSpaceName: .deviceRGB, bytesPerRow: 0, bitsPerPixel: 0)!
NSGraphicsContext.saveGraphicsState()
NSGraphicsContext.current = NSGraphicsContext(bitmapImageRep: rep)
let ink = NSColor(srgbRed: 48/255, green: 52/255, blue: 59/255, alpha: 1)
let cyan = NSColor(srgbRed: 35/255, green: 186/255, blue: 236/255, alpha: 1)
// "A" (top-left), drawn as strokes so it stays crisp at 16 px
let a = NSBezierPath()
a.lineWidth = 1.6 * s
a.lineCapStyle = .butt
a.lineJoinStyle = .miter
a.move(to: NSPoint(x: 0.8 * s, y: 5.6 * s))
a.line(to: NSPoint(x: 4.6 * s, y: 15.2 * s))
a.line(to: NSPoint(x: 8.4 * s, y: 5.6 * s))
a.move(to: NSPoint(x: 2.3 * s, y: 9.0 * s))
a.line(to: NSPoint(x: 6.9 * s, y: 9.0 * s))
ink.setStroke(); a.stroke()
// cyan replace arrow (bottom-right): down from the A, then right, ending in an arrow head
let arrow = NSBezierPath()
arrow.lineWidth = 1.6 * s
arrow.move(to: NSPoint(x: 10.0 * s, y: 13.0 * s))
arrow.line(to: NSPoint(x: 10.0 * s, y: 4.0 * s))
arrow.line(to: NSPoint(x: 12.4 * s, y: 4.0 * s))
cyan.setStroke(); arrow.stroke()
let head = NSBezierPath()
head.move(to: NSPoint(x: 12.0 * s, y: 7.0 * s))
head.line(to: NSPoint(x: 15.6 * s, y: 4.0 * s))
head.line(to: NSPoint(x: 12.0 * s, y: 1.0 * s))
head.close()
cyan.setFill(); head.fill()
NSGraphicsContext.restoreGraphicsState()
try! rep.representation(using: .png, properties: [:])!.write(to: URL(fileURLWithPath: out))
