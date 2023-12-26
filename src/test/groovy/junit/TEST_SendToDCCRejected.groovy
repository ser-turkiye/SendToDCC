package junit

import de.ser.doxis4.agentserver.AgentExecutionResult
import org.junit.*
import ser.SendToDCCRejected

class TEST_SendToDCCRejected {

    Binding binding

    @BeforeClass
    static void initSessionPool() {
        AgentTester.initSessionPool()
    }

    @Before
    void retrieveBinding() {
        binding = AgentTester.retrieveBinding()
    }

    @Test
    void testForAgentResult() {
        def agent = new SendToDCCRejected();

        binding["AGENT_EVENT_OBJECT_CLIENT_ID"] = "ST03BPM24cc2f2d69-87b6-4e3b-8434-c740b1349f40182023-12-26T08:07:27.676Z013"

        def result = (AgentExecutionResult) agent.execute(binding.variables)
        assert result.resultCode == 0
    }

    @Test
    void testForJavaAgentMethod() {
        //def agent = new JavaAgent()
        //agent.initializeGroovyBlueline(binding.variables)
        //assert agent.getServerVersion().contains("Linux")
    }

    @After
    void releaseBinding() {
        AgentTester.releaseBinding(binding)
    }

    @AfterClass
    static void closeSessionPool() {
        AgentTester.closeSessionPool()
    }
}
