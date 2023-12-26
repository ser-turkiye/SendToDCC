package junit

import de.ser.doxis4.agentserver.AgentExecutionResult
import org.junit.*
import ser.SendToDCCApproved

class TEST_SendToDCCApproved {

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
        def agent = new SendToDCCApproved();

        binding["AGENT_EVENT_OBJECT_CLIENT_ID"] = "ST03BPM244bb35110-cb3c-4e73-b464-ec0a50665759182023-12-25T10:00:18.569Z014"

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
