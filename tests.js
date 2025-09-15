/**
 * Web app entry point for running tests and viewing results in a browser.
 */
function doGet(e) {
  QUnitGS2.init();

  runAllTests();

  QUnitGS2.QUnit.start();
  return QUnitGS2.getHtml();
}

/**
 * Retrieve test results when ready.
 * This is called by the HTML front-end to poll for results.
 */
function getResultsFromServer() {
   return QUnitGS2.getResultsFromServer();
}

function runAllTests() {
  testItemMiscCost();
  // Other tests can be called here in the future.
}

function testItemMiscCost() {
  QUnitGS2.QUnit.module('itemMiscCost - Calculation Test');

  QUnitGS2.QUnit.test('Calculations for different item types', function(assert) {
    // Test Case 1: Charge_Fee
    var item1 = [1, new Date(), "PropertyA", 1000, "0.Charge_Fee", "note", "", null, null, null];
    var miscCost1 = new itemMiscCost(item1);
    assert.equal(miscCost1.expect_misc(), 1000, "expect_misc() for Charge_Fee should be positive");
    assert.equal(miscCost1.balance_misc(), 0, "balance_misc() for Charge_Fee should be zero");

    // Test Case 2: Cash_Rent
    var item2 = [2, new Date(), "PropertyB", 1500, "1.Cash_Rent", "note", "", null, null, null];
    var miscCost2 = new itemMiscCost(item2);
    assert.equal(miscCost2.expect_misc(), -1500, "expect_misc() for Cash_Rent should be negative");
    assert.equal(miscCost2.balance_misc(), 1500, "balance_misc() for Cash_Rent should be positive");

    // Test Case 3: Commission
    var item3 = [3, new Date(), "PropertyC", 500, "3.Commission", "note", "", null, null, null];
    var miscCost3 = new itemMiscCost(item3);
    assert.equal(miscCost3.expect_misc(), -500, "expect_misc() for Commission should be negative");
    assert.equal(miscCost3.balance_misc(), -500, "balance_misc() for Commission should be negative");

    // Test Case 4: Refund
    var item4 = [4, new Date(), "PropertyD", 200, "6.Refund", "note", "", null, null, null];
    var miscCost4 = new itemMiscCost(item4);
    assert.equal(miscCost4.expect_misc(), -200, "expect_misc() for Refund should be negative");
    assert.equal(miscCost4.balance_misc(), -200, "balance_misc() for Refund should be negative");
  });
}