/*!
 * Copyright 2024 Thai Pangsakulyanont (JavaScript Port)
 * Copyright 2022 Thomas Reidemeister (Original Python Code)
 *
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 *
 *     http://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 *
 * This project is a JavaScript port of the brother_pt project
 * (https://github.com/treideme/brother_pt) by Thomas Reidemeister,
 * which is licensed under the Apache License, Version 2.0,
 * and adapted for WebUSB use.
 */

// @ts-check

/**
 * Brother P-Touch printer client for Web USB.
 * Tested with Brother P-Touch Cube (PT-P710BT) with a 18mm tape.
 */

// Raster command reference:
// https://download.brother.com/welcome/docp100064/cv_pte550wp750wp710bt_eng_raster_102.pdf

const PRINT_HEAD_PINS = 128;
const LINE_LENGTH_BYTES = 16;

// See table on page 14 of the raster command reference
export const TAPE_MARGINS = {
  4: 52, // 3.5mm
  6: 48, // 6mm
  9: 39, // 9mm
  12: 29, // 12mm
  18: 8, // 19mm
  24: 0, // 24mm
};

/**
 * @param {number} mediaWidth
 */
export function getPrintWidth(mediaWidth) {
  const margin = TAPE_MARGINS[mediaWidth];
  if (margin == null) {
    throw new Error(`Unsupported media width: ${mediaWidth}`);
  }
  return PRINT_HEAD_PINS - TAPE_MARGINS[mediaWidth] * 2;
}

/**
 * USB vendor ID for Brother printers
 */
export const VENDOR_ID = 0x04f9;

/**
 * Map of supported printer product IDs
 */
export const SUPPORTED_PRINTER_IDS = {
  E550W: 0x2060,
  P750W: 0x2062,
  P710BT: 0x20af,
};

export const STATUS_MESSAGE_LENGTH = 32;

// Status offsets enum from cmd.py
export const StatusOffsets = {
  ERROR_INFORMATION_1: 8,
  ERROR_INFORMATION_2: 9,
  MEDIA_WIDTH: 10,
  MEDIA_TYPE: 11,
  MODE: 15,
  MEDIA_LENGTH: 17,
  STATUS_TYPE: 18,
  PHASE_TYPE: 19,
  PHASE_NUMBER: 20,
  NOTIFICATION_NUMBER: 22,
  TAPE_COLOR_INFORMATION: 24,
  TEXT_COLOR_INFORMATION: 25,
};

const MINIMUM_TAPE_POINTS = 174; // 25.4 mm @ 180dpi

/**
 * Class representing a Brother P-Touch printer connected via WebUSB
 */
export class BrotherPrinter {
  constructor() {
    this.device = null;
    this.outEndpoint = null;
    this.inEndpoint = null;
    this._mediaWidth = null;
    this._mediaType = null;
    this._tapeColor = null;
    this._textColor = null;

    /** @type {(status: any) => void} */
    this.onstatus = () => {};
  }

  /**
   * Connects to a Brother P-Touch printer via WebUSB and initializes it
   * @throws {Error} If no printer is found or connection fails
   */
  async connect() {
    this.device = await navigator.usb.requestDevice({
      filters: [{ vendorId: VENDOR_ID }],
    });

    await this.device.open();
    await this.device.selectConfiguration(1);

    const usbInterface = this.device.configuration.interfaces[0];
    const alternate = usbInterface.alternates[0];

    await this.device.claimInterface(usbInterface.interfaceNumber);
    await this.device.selectAlternateInterface(usbInterface.interfaceNumber, 0);

    this.outEndpoint = alternate.endpoints.find((e) => e.direction === "out");
    this.inEndpoint = alternate.endpoints.find((e) => e.direction === "in");

    await this._updateStatus();
  }

  /**
   * Writes data to the printer
   * @param {Uint8Array} data - The data to write
   * @private
   */
  async write(data) {
    await this.device.transferOut(this.outEndpoint.endpointNumber, data);
  }

  /**
   * Reads data from the printer
   * @param {number} length - Number of bytes to read (default: 32)
   * @returns {Promise<Uint8Array>} The data read from the printer
   * @private
   */
  async read(length = STATUS_MESSAGE_LENGTH) {
    const result = await this.device.transferIn(
      this.inEndpoint.endpointNumber,
      length
    );
    return new Uint8Array(result.data.buffer);
  }

  /**
   * Updates the printer status by requesting and reading current state
   */
  async _updateStatus() {
    // Initialize printer
    await this.write(new Uint8Array(100).fill(0)); // invalidate
    await this.write(new Uint8Array([0x1b, 0x40])); // initialize

    // Request status
    await this.write(new Uint8Array([0x1b, 0x69, 0x53]));
    const data = await this.read();

    // Update printer state
    this._mediaWidth = data[StatusOffsets.MEDIA_WIDTH];
    this._mediaType = data[StatusOffsets.MEDIA_TYPE];
    this._tapeColor = data[StatusOffsets.TAPE_COLOR_INFORMATION];
    this._textColor = data[StatusOffsets.TEXT_COLOR_INFORMATION];
    this._notifyStatus(data);
  }

  /**
   * @param {Uint8Array} data
   */
  _notifyStatus(data) {
    this.onstatus({
      /** The first error status byte. Bit flags indicate various error conditions:
       * - 0x01: No media
       * - 0x04: Cutter jam
       * - 0x08: Low batteries
       * - 0x40: High-voltage adapter issue
       */
      errorInfo1: data[StatusOffsets.ERROR_INFORMATION_1],
      /** The second error status byte. Bit flags indicate various error conditions:
       * - 0x01: Wrong media size
       * - 0x10: Cover open
       * - 0x20: Overheating
       */
      errorInfo2: data[StatusOffsets.ERROR_INFORMATION_2],
      /** The width of the tape media currently loaded in the printer (in mm) */
      mediaWidth: this._mediaWidth,
      /** The type code indicating the kind of tape media loaded in the printer */
      mediaType: this._mediaType,
      /** The current operational mode of the printer */
      mode: data[StatusOffsets.MODE],
      /** The length of the tape media (in mm) */
      mediaLength: data[StatusOffsets.MEDIA_LENGTH],
      /** The current status type of the printer:
       * - 0x00: Reply to status request
       * - 0x01: Printing completed
       * - 0x02: Error occurred
       */
      statusType: data[StatusOffsets.STATUS_TYPE],
      /** The current phase type of the printer operation */
      phaseType: data[StatusOffsets.PHASE_TYPE],
      /** The current phase number of the printer operation */
      phaseNumber: data[StatusOffsets.PHASE_NUMBER],
      /** The notification number for printer status updates */
      notificationNumber: data[StatusOffsets.NOTIFICATION_NUMBER],
      /** The color code of the tape currently loaded in the printer */
      tapeColor: this._tapeColor,
      /** The color code for the text/print color of the current tape */
      textColor: this._textColor,
    });
  }

  /**
   * Gets the width of the tape media currently loaded in the printer
   * @returns {number} The media width in millimeters
   */
  get mediaWidth() {
    if (this._mediaWidth == null) {
      throw new Error("Printer not initialized or media width is unknown");
    }
    return this._mediaWidth;
  }

  /**
   * Gets the type of the tape media currently loaded in the printer
   * @returns {number} The media type code
   */
  get mediaType() {
    if (this._mediaType == null) {
      throw new Error("Printer not initialized or media type is unknown");
    }
    return this._mediaType;
  }

  /**
   * Gets the color of the tape currently loaded in the printer
   * @returns {number} The tape color code
   */
  get tapeColor() {
    if (this._tapeColor == null) {
      throw new Error("Printer not initialized or tape color is unknown");
    }
    return this._tapeColor;
  }

  /**
   * Gets the text color of the tape currently loaded in the printer
   * @returns {number} The text color code
   */
  get textColor() {
    if (this._textColor == null) {
      throw new Error("Printer not initialized or text color is unknown");
    }
    return this._textColor;
  }

  /**
   * Returns the number of pixels that can be printed on a single column of tape.
   */
  get printWidth() {
    if (!this.mediaWidth) {
      throw new Error("Printer not initialized or media width is unknown");
    }
    return getPrintWidth(this.mediaWidth);
  }

  /**
   * Returns the minimum number of columns required to print a tape.
   */
  get minimumColumns() {
    return MINIMUM_TAPE_POINTS;
  }

  /**
   * Prints an image.
   * @param {(0 | 255)[][]} columns - The columns to print.
   *  This should be a 2D array of pixel values (0 or 255) where each sub-array represents a column of pixels.
   *  There must be at least 174 columns (because each print must be at least 1 inch long).
   *  In each column must have the same number of pixels as the printWidth (alternatively, you can use 128 pixels if you want to print beyond the tape width, for experimental purposes).
   *  0 results in the tape color, while 255 results in a text color.
   * @param {boolean} [isLastPage=true] - Whether this is the last page to print.
   *  If false, the printer will not cut the tape after printing, so the last page will still be stuck to the printer. It will be cut when the next page is printed.
   *  If true, the printer will cut the tape right after printing. You will get the tape, but on the next print some tape will be wasted.
   * @throws {Error} If a printer error occurs during printing (e.g., no media, cover open, etc.)
   */
  async print(columns, isLastPage = true) {
    if (columns.length < MINIMUM_TAPE_POINTS) {
      throw new Error(
        `Image is too short: ${columns.length} < ${MINIMUM_TAPE_POINTS}`
      );
    }
    const printWidth = this.printWidth;
    const margin = (128 - printWidth) / 2;
    const raster = columns.map((column) => {
      if (column.length !== printWidth && column.length !== 128) {
        throw new Error(
          `Column width ${column.length} does not match print width ${printWidth}`
        );
      }
      return columnToRaster(column, column.length === 128 ? 0 : margin);
    });

    // Enter dynamic command mode
    await this.write(new Uint8Array([0x1b, 0x69, 0x61, 0x01]));

    // Enable status notification
    await this.write(new Uint8Array([0x1b, 0x69, 0x21, 0x00]));

    // Print information command
    const numberOfBytes = raster.reduce((sum, chunk) => sum + chunk.length, 0);
    const dataLength = numberOfBytes >> 4;
    const printInfoCmd = new Uint8Array([
      0x1b,
      0x69,
      0x7a,
      0x84,
      0x00,
      this.mediaWidth,
      0x00,
      // Data length (4 bytes little endian)
      (dataLength >> 0) & 0xff,
      (dataLength >> 8) & 0xff,
      (dataLength >> 16) & 0xff,
      (dataLength >> 24) & 0xff,
      0x00,
      0x00,
    ]);
    await this.write(printInfoCmd);

    // Set mode (auto-cut, no mirror)
    await this.write(new Uint8Array([0x1b, 0x69, 0x4d, 0x40]));

    // Set advanced mode
    await this.write(new Uint8Array([0x1b, 0x69, 0x4b, 0x08]));

    // Set margin to 0
    await this.write(new Uint8Array([0x1b, 0x69, 0x64, 0x00, 0x00]));

    // Set no compression mode
    await this.write(new Uint8Array([0x4d, 0x00]));

    // Send raster data in 16-byte chunks
    for (const chunk of raster) {
      const cmd = new Uint8Array([
        0x47, // Raster command
        chunk.length,
        0x00, // Length (little endian)
        ...chunk,
      ]);
      await this.write(cmd);
    }

    // Print and feed (or print without feeding)
    await this.write(new Uint8Array([isLastPage ? 0x1a : 0x0c]));

    // Wait for print to complete
    while (true) {
      const status = await this.read();
      this._notifyStatus(status);

      // Check status type
      if (status[18] === 0x01) {
        // Print completed
        this._notifyStatus(await this.read()); // Absorb phase change message
        break;
      } else if (status[18] === 0x02) {
        // Error occurred
        let errorMessage = [];

        // Error 1
        if (status[8]) {
          if (status[8] & 0x01) errorMessage.push("no media");
          if (status[8] & 0x04) errorMessage.push("cutter jam");
          if (status[8] & 0x08) errorMessage.push("low batteries");
          if (status[8] & 0x40) errorMessage.push("high-voltage adapter");
        }

        // Error 2
        if (status[9]) {
          if (status[9] & 0x01) errorMessage.push("wrong media (check size)");
          if (status[9] & 0x10) errorMessage.push("cover open");
          if (status[9] & 0x20) errorMessage.push("overheating");
        }

        throw new Error(errorMessage.join("|"));
      }
    }
  }
}

/**
 * @param {(0 | 255)[]} column
 * @param {number} margin
 */
function columnToRaster(column, margin) {
  const lineLength = LINE_LENGTH_BYTES;
  const buffer = new Uint8Array(lineLength);
  const columnBits = [];

  // Leading margin
  for (let i = 0; i < margin; i++) {
    columnBits.push(0);
  }

  // Image data
  for (const pixel of column) {
    columnBits.push(pixel ? 1 : 0);
  }

  // Trailing margin
  while (columnBits.length < lineLength * 8) {
    columnBits.push(0);
  }

  // Pack bits into bytes
  for (let byteIndex = 0; byteIndex < lineLength; byteIndex++) {
    let byte = 0;
    for (let bit = 0; bit < 8; bit++) {
      const bitIndex = byteIndex * 8 + bit;
      if (columnBits[bitIndex]) {
        byte |= 1 << (7 - bit);
      }
    }
    buffer[byteIndex] = byte;
  }

  return buffer;
}
