package driverFactory;

import org.testng.annotations.Test;

public class AppTest extends DriverScript {
@Test
public void kickStart() throws Throwable {
	DriverScript ds=new DriverScript();
	ds.start_test();
}
}
