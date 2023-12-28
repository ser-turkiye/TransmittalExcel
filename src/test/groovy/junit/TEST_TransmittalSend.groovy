package junit

import de.ser.doxis4.agentserver.AgentExecutionResult
import org.junit.*
import ser.TransmittalSend

class TEST_TransmittalSend {

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
        def agent = new TransmittalSend();

        binding["AGENT_EVENT_OBJECT_CLIENT_ID"] = "ST03BPM249601d8ad-ec15-4bb5-8425-ef07c33456f1182023-12-28T08:49:09.349Z014"

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
