package junit

import de.ser.doxis4.agentserver.AgentExecutionResult
import org.junit.*
import ser.TransmittalFromExcel

class TEST_TransmittalFromExcel {

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
        def agent = new TransmittalFromExcel();

        binding["AGENT_EVENT_OBJECT_CLIENT_ID"] = "SD06D_QCON24c6b44b84-aeb3-4e85-9798-5c6e1998b4d0182023-12-15T17:38:33.553Z011"

        def result = (AgentExecutionResult)agent.execute(binding.variables)
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
