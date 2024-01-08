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

        binding["AGENT_EVENT_OBJECT_CLIENT_ID"] = "SP03BPM240dd56baa-4105-4e5e-aad9-7957c4a74925182024-01-08T07:17:28.357Z00"

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
