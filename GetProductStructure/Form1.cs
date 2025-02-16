using System;                                            // 기본 .NET 클래스들을 포함
using System.Runtime.InteropServices;                    // COM 관련 기능 (Marshal 클래스 등)을 사용하기 위함
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;                              // Windows Forms 관련 클래스들을 포함
using INFITF;                                            // CATIA 애플리케이션 관련 인터페이스를 포함 (예: Application)
using MECMOD;                                            // CATIA 모델 관련 인터페이스를 포함
using ProductStructureTypeLib;                           // CATIA 어셈블리(제품 구조) 관련 인터페이스를 포함

namespace GetProductStructure
{
    // Windows Form 클래스 정의
    public partial class Form1 : Form
    {
        // CATIA 애플리케이션 객체를 저장할 변수 선언
        INFITF.Application myCATIA;

        // 폼 생성자: 폼의 구성요소들을 초기화
        public Form1()
        {
            InitializeComponent(); // 디자이너에서 설정한 컨트롤 및 레이아웃 초기화
        }

        // 폼이 로드될 때 호출되는 이벤트 핸들러
        private void Form1_Load(object sender, EventArgs e)
        {
            // 이미 실행 중인 CATIA 애플리케이션에 연결
            myCATIA = (INFITF.Application)Marshal.GetActiveObject("CATIA.Application");
            Console.WriteLine("Checked. Connected."); // 연결 성공 메시지 출력
        }

        // 버튼 클릭 시 호출되는 이벤트 핸들러
        private void button1_Click(object sender, EventArgs e)
        {
            // 폼 로드 이벤트 핸들러를 호출하여 CATIA 연결을 보장 (일부 상황에서는 로드가 제대로 호출되지 않을 수 있음)
            Form1_Load(sender, e);

            // CATIA에서 현재 활성 문서를 가져옴
            Document oDoc = myCATIA.ActiveDocument;
            // 활성 문서를 ProductDocument(어셈블리 문서)로 캐스팅
            ProductDocument productDocument1 = (ProductDocument)oDoc;
            // 어셈블리의 루트 Product의 이름을 가져와 TreeView의 루트 노드를 생성
            TreeNode rootNode1 = new TreeNode() { Text = productDocument1.Product.get_Name() };

            // TreeView 컨트롤에 루트 노드를 추가
            treeView1.Nodes.Add(rootNode1);
            // 루트 Product의 하위 제품 구조를 재귀적으로 탐색하여 TreeView에 추가
            GetProductStructure(productDocument1.Product, rootNode1);

            Console.WriteLine("Finish!"); // 처리 완료 메시지 출력
        }

        // 어셈블리 제품 구조를 재귀적으로 탐색하는 메서드
        // oInstance: 현재 탐색할 Product 객체, treeNode1: 해당 제품을 나타내는 TreeView 노드
        public void GetProductStructure(Product oInstance, TreeNode treeNode1)
        {
            // 현재 Product 객체의 자식 제품(어셈블리 내의 부품) 컬렉션을 가져옴
            Products oInstances = oInstance.Products;
            // 자식 제품의 총 개수를 확인
            int numberOfInstance = oInstances.Count;

            // 전달된 TreeView 노드가 null이면 종료
            if (treeNode1 == null)
            {
                return;
            }

            // 자식 제품이 하나도 없으면 종료
            if (numberOfInstance == 0)
            {
                return;
            }

            // 자식 제품이 존재할 경우, 각 제품에 대해 처리
            if (numberOfInstance > 0)
            {
                // CATIA COM 객체의 인덱스는 1부터 시작하므로 1부터 반복
                for (int i = 1; i <= numberOfInstance; i++)
                {
                    // 현재 자식 제품의 이름을 가져와 새로운 TreeView 노드를 생성
                    TreeNode iNode = new TreeNode() { Text = oInstances.Item(i).get_Name() };
                    // 생성된 노드를 부모 TreeView 노드에 추가
                    treeNode1.Nodes.Add(iNode);
                    // 재귀 호출로 현재 자식 제품의 하위 제품 구조를 탐색하여 TreeView에 추가
                    GetProductStructure(oInstances.Item(i), iNode);
                }
            }
        }
    }
}
