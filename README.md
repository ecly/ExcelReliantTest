# How to test:
**Prerequisites**  
Identical or simliar installations as seen in ***Versions***

**Versions**  
Windows: 7  
Visual Studio: 2017 Enterprise  
.NET: 4.5.2  
NUnit: 2.6.4  
NUnit Adapter: NUnit 2 Test Adapter 2.1.1  
Excel: Exel 2016

### Reproduction:
- Open solution and build to discover unit tests.  
- Run both the tests without Code Coverage. *(NUnit 2 Test Adapter needed for NUnit test)*
 + Observe them both executing just fine.
- Run one of them with Code Coverage.
 + Hangs indefinitely and never closes instance.
- To prepare for second test, restart Visual Studio and ensure that the background Excel instance is closed manually. Do this through Task Manager and merely look for 'EXCEL.EXE'.
 - Now attempt to run the second test and observe the same behavior as seen in first test.
