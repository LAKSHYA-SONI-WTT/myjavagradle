package com.project.runner;

import org.testng.annotations.Listeners; 
import org.testng.annotations.Test;
import com.aventstack.extentreports.testng.listener.ExtentITestListenerClassAdapter;
import io.cucumber.junit.CucumberOptions;
import io.cucumber.testng.AbstractTestNGCucumberTests;

@CucumberOptions(
    features = "src/test/resources/features",
    glue = "stepDefinitions",
    publish=true,
    plugin={"pretty","html:target/CucumberReports/CucumberReport.html"},
    monochrome=true,
    tags = "@test"
)


@Listeners({ExtentITestListenerClassAdapter.class})
public class TestRunner extends AbstractTestNGCucumberTests {

    @Test(priority = 1)
    public void runDistrictFeature() {
        runCucumberFeature("src/test/resources/com/project/features/test.feature");
    }
}
