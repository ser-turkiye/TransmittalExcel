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

        binding["AGENT_EVENT_OBJECT_CLIENT_ID"] = "ST03BPM240193a046-5b35-47ca-b9d3-076a6ba85099182024-01-02T14:30:27.549Z012"

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
