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

        binding["AGENT_EVENT_OBJECT_CLIENT_ID"] = "ST03BPM24d98cc4b2-d187-46f9-899e-f96892a618e5182023-12-29T09:55:58.992Z014"

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
