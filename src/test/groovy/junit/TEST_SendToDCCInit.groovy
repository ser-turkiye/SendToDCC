package junit

import de.ser.doxis4.agentserver.AgentExecutionResult
import org.junit.*
import ser.SendToDCCInit

class TEST_SendToDCCInit {

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
        def agent = new SendToDCCInit();

        binding["AGENT_EVENT_OBJECT_CLIENT_ID"] = "SP03BPM245c3ff3f3-6d0e-4e1a-b424-b6f4f8788146182024-01-02T08:08:46.712Z00"

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
