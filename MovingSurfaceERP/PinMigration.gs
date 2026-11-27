/** Run once after deploying this upgrade to preserve leading zeroes in existing PINs. */
function setupFourDigitPins() {
  requireOwnerContext_();
  migratePinsToFourDigits();
  Logger.log('PIN migration complete. PINs are now stored as four-digit text values.');
}
