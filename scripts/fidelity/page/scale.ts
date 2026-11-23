// The scale at which the harness paints a fixture: CSS pixels for each point on the harness page,
// and device pixels for each CSS pixel in the screenshot that the pixel score compares
// (`run.ts`). The print paint for the line scores uses one device pixel for each CSS pixel
// (`page/main.ts`). Tests that rebuild what the harness paints import these, so they cannot
// drift from the harness.

/** CSS pixels for each point on the harness page. */
export const PX_PER_PT = 4 / 3;
/** Device pixels for each CSS pixel in the screenshot of the harness page. */
export const DEVICE_SCALE = 3;
