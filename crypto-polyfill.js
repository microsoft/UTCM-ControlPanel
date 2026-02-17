// Polyfill for crypto.subtle.digest in non-secure contexts (HTTP).
// crypto.subtle is only available in secure contexts (HTTPS/localhost).
// MSAL requires crypto.subtle.digest('SHA-256', ...) for PKCE generation.
// This provides a pure-JS SHA-256 fallback when the native API is unavailable.
(function () {
  if (typeof window === 'undefined') return;
  if (!window.crypto) {
    window.crypto = {};
  }
  // If crypto.subtle is already available, no polyfill needed
  if (window.crypto.subtle) return;

  // SHA-256 round constants
  var K = [
    0x428a2f98, 0x71374491, 0xb5c0fbcf, 0xe9b5dba5,
    0x3956c25b, 0x59f111f1, 0x923f82a4, 0xab1c5ed5,
    0xd807aa98, 0x12835b01, 0x243185be, 0x550c7dc3,
    0x72be5d74, 0x80deb1fe, 0x9bdc06a7, 0xc19bf174,
    0xe49b69c1, 0xefbe4786, 0x0fc19dc6, 0x240ca1cc,
    0x2de92c6f, 0x4a7484aa, 0x5cb0a9dc, 0x76f988da,
    0x983e5152, 0xa831c66d, 0xb00327c8, 0xbf597fc7,
    0xc6e00bf3, 0xd5a79147, 0x06ca6351, 0x14292967,
    0x27b70a85, 0x2e1b2138, 0x4d2c6dfc, 0x53380d13,
    0x650a7354, 0x766a0abb, 0x81c2c92e, 0x92722c85,
    0xa2bfe8a1, 0xa81a664b, 0xc24b8b70, 0xc76c51a3,
    0xd192e819, 0xd6990624, 0xf40e3585, 0x106aa070,
    0x19a4c116, 0x1e376c08, 0x2748774c, 0x34b0bcb5,
    0x391c0cb3, 0x4ed8aa4a, 0x5b9cca4f, 0x682e6ff3,
    0x748f82ee, 0x78a5636f, 0x84c87814, 0x8cc70208,
    0x90befffa, 0xa4506ceb, 0xbef9a3f7, 0xc67178f2
  ];

  function rightRotate(value, amount) {
    return (value >>> amount) | (value << (32 - amount));
  }

  function sha256(data) {
    var bytes = new Uint8Array(data);
    var msgLen = bytes.length;
    var bitLen = msgLen * 8;

    // Pad message to multiple of 64 bytes
    var paddedLen = Math.ceil((msgLen + 9) / 64) * 64;
    var padded = new Uint8Array(paddedLen);
    padded.set(bytes);
    padded[msgLen] = 0x80;

    // Append original length in bits as 64-bit big-endian
    var view = new DataView(padded.buffer);
    view.setUint32(paddedLen - 4, bitLen >>> 0, false);
    view.setUint32(paddedLen - 8, (bitLen / 0x100000000) >>> 0, false);

    // Initial hash values
    var h0 = 0x6a09e667, h1 = 0xbb67ae85, h2 = 0x3c6ef372, h3 = 0xa54ff53a;
    var h4 = 0x510e527f, h5 = 0x9b05688c, h6 = 0x1f83d9ab, h7 = 0x5be0cd19;

    var w = new Array(64);

    // Process each 64-byte block
    for (var offset = 0; offset < paddedLen; offset += 64) {
      for (var i = 0; i < 16; i++) {
        w[i] = view.getUint32(offset + i * 4, false);
      }

      for (var i = 16; i < 64; i++) {
        var s0 = rightRotate(w[i - 15], 7) ^ rightRotate(w[i - 15], 18) ^ (w[i - 15] >>> 3);
        var s1 = rightRotate(w[i - 2], 17) ^ rightRotate(w[i - 2], 19) ^ (w[i - 2] >>> 10);
        w[i] = (w[i - 16] + s0 + w[i - 7] + s1) | 0;
      }

      var a = h0, b = h1, c = h2, d = h3;
      var e = h4, f = h5, g = h6, h = h7;

      for (var i = 0; i < 64; i++) {
        var S1 = rightRotate(e, 6) ^ rightRotate(e, 11) ^ rightRotate(e, 25);
        var ch = (e & f) ^ (~e & g);
        var temp1 = (h + S1 + ch + K[i] + w[i]) | 0;
        var S0 = rightRotate(a, 2) ^ rightRotate(a, 13) ^ rightRotate(a, 22);
        var maj = (a & b) ^ (a & c) ^ (b & c);
        var temp2 = (S0 + maj) | 0;

        h = g;
        g = f;
        f = e;
        e = (d + temp1) | 0;
        d = c;
        c = b;
        b = a;
        a = (temp1 + temp2) | 0;
      }

      h0 = (h0 + a) | 0;
      h1 = (h1 + b) | 0;
      h2 = (h2 + c) | 0;
      h3 = (h3 + d) | 0;
      h4 = (h4 + e) | 0;
      h5 = (h5 + f) | 0;
      h6 = (h6 + g) | 0;
      h7 = (h7 + h) | 0;
    }

    var result = new ArrayBuffer(32);
    var resultView = new DataView(result);
    resultView.setInt32(0, h0, false);
    resultView.setInt32(4, h1, false);
    resultView.setInt32(8, h2, false);
    resultView.setInt32(12, h3, false);
    resultView.setInt32(16, h4, false);
    resultView.setInt32(20, h5, false);
    resultView.setInt32(24, h6, false);
    resultView.setInt32(28, h7, false);

    return result;
  }

  // Polyfill crypto.subtle with SHA-256 digest support
  window.crypto.subtle = {
    digest: function (algorithm, data) {
      var algo = typeof algorithm === 'string' ? algorithm : algorithm.name;
      if (algo === 'SHA-256') {
        try {
          return Promise.resolve(sha256(data));
        } catch (e) {
          return Promise.reject(e);
        }
      }
      return Promise.reject(new Error('crypto.subtle polyfill: unsupported algorithm ' + algo));
    }
  };

  console.warn(
    'crypto.subtle is not natively available (non-secure context). ' +
    'Using SHA-256 polyfill for PKCE. For best security, serve this app over HTTPS.'
  );
})();
