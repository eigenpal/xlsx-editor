// The placement of the menu bar's panels (`menuPanelPlacement`): a bar menu's panel and a
// submenu stay inside the viewport, a submenu opens toward the start of the bar when the end has
// no room, and a panel is limited to the viewport height.

import { describe, expect, test } from 'bun:test';
import { Window } from 'happy-dom';
import { menuPanelPlacement, placeMenuPanel } from '../menu-placement.ts';

describe('menu panel placement', () => {
  const view = { viewportWidth: 400, viewportHeight: 300 };
  const box = (left: number, right: number, top = 0, bottom = 30) => ({ left, right, top, bottom });

  test('a bar panel lines up with its trigger start, and with its end when the start has no room', () => {
    const at = (left: number, right: number, rtl = false) =>
      menuPanelPlacement({
        submenu: false,
        rtl,
        anchor: box(left, right),
        width: 240,
        height: 200,
        ...view,
      });
    expect(at(20, 60)).toMatchObject({ left: 20, top: 32, side: 'start' });
    expect(at(300, 340)).toMatchObject({ left: 100, top: 32, side: 'end' });
    // Right to left: the start edge is the right edge.
    expect(at(300, 340, true)).toMatchObject({ left: 100, top: 32, side: 'start' });
    expect(at(20, 60, true)).toMatchObject({ left: 20, top: 32, side: 'end' });
  });

  test('the More panel lines up with its end first', () => {
    const at = menuPanelPlacement({
      submenu: false,
      rtl: false,
      prefer: 'end',
      anchor: box(300, 340),
      width: 240,
      height: 200,
      ...view,
    });
    expect(at).toMatchObject({ left: 100, top: 32, side: 'end' });
  });

  test('a panel wider than the room is clamped to the margins', () => {
    const at = menuPanelPlacement({
      submenu: false,
      rtl: false,
      anchor: box(150, 190),
      width: 390,
      height: 100,
      ...view,
    });
    expect(at).toMatchObject({ left: 8, top: 32, side: 'clamped' });
  });

  test('a submenu opens at the end of its row, toward the start when the end has no room', () => {
    const at = (left: number, right: number, rtl = false) =>
      menuPanelPlacement({
        submenu: true,
        rtl,
        anchor: box(left, right, 100, 130),
        width: 150,
        height: 100,
        ...view,
      });
    expect(at(0, 200)).toMatchObject({ left: 204, top: 95, side: 'end' });
    expect(at(150, 350)).toMatchObject({ left: 242, top: 95, side: 'clamped' });
    expect(at(170, 390)).toMatchObject({ left: 16, top: 95, side: 'start' });
    expect(at(200, 390, true)).toMatchObject({ left: 46, top: 95, side: 'end' });
    expect(at(10, 200, true)).toMatchObject({ left: 204, top: 95, side: 'start' });
  });

  test('a bar panel is limited to the viewport space below its trigger, a submenu to the viewport', () => {
    const bar = menuPanelPlacement({
      submenu: false,
      rtl: false,
      anchor: box(20, 60, 0, 30),
      width: 240,
      height: 900,
      ...view,
    });
    // 300 high, an 8 px margin, and a top at 32: 260 px of room.
    expect(bar.maxHeight).toBe(260);
    const sub = menuPanelPlacement({
      submenu: true,
      rtl: false,
      anchor: box(0, 100, 50, 80),
      width: 150,
      height: 900,
      ...view,
    });
    expect(sub).toMatchObject({ top: 8, maxHeight: 284 });
    const none = menuPanelPlacement({
      submenu: false,
      rtl: false,
      anchor: box(20, 60, 290, 320),
      width: 240,
      height: 100,
      ...view,
    });
    expect(none.maxHeight).toBe(0);
  });

  test('a submenu moves up to keep its bottom inside the viewport', () => {
    const at = menuPanelPlacement({
      submenu: true,
      rtl: false,
      anchor: box(0, 100, 250, 280),
      width: 150,
      height: 120,
      ...view,
    });
    expect(at.top).toBe(172);
  });
});

describe('placeMenuPanel', () => {
  test('a submenu of a panel that scrolls stands against the viewport and follows its row', () => {
    const win = new Window({ width: 800, height: 400 });
    const doc = win.document as unknown as Document;
    const outer = doc.createElement('div');
    outer.className = 'xlsx-menu-panel';
    outer.setAttribute('data-scrolls', '');
    const sub = doc.createElement('div');
    const row = doc.createElement('button');
    row.setAttribute('aria-haspopup', 'menu');
    const panel = doc.createElement('div');
    sub.append(row, panel);
    outer.append(sub);
    doc.body.append(outer);
    let rowTop = 100;
    const rect = (top: number) =>
      ({ left: 100, right: 300, top, bottom: top + 30, width: 200, height: 30 }) as DOMRect;
    row.getBoundingClientRect = () => rect(rowTop);
    sub.getBoundingClientRect = () => rect(rowTop);
    // The panel box follows the inline place: no transformed host here.
    panel.getBoundingClientRect = () => rect(parseFloat(panel.style.top) || 0);
    Object.defineProperty(panel, 'offsetWidth', { value: 150 });
    Object.defineProperty(panel, 'offsetHeight', { value: 100 });
    const stop = placeMenuPanel(panel, { submenu: true });
    expect(panel.style.position).toBe('fixed');
    expect(panel.style.top).toBe('95px');
    rowTop = 60;
    outer.dispatchEvent(new win.Event('scroll') as unknown as Event);
    expect(panel.style.top).toBe('55px');
    stop();
    expect(panel.style.top).toBe('');
    rowTop = 20;
    outer.dispatchEvent(new win.Event('scroll') as unknown as Event);
    expect(panel.style.top).toBe('');
  });
});
